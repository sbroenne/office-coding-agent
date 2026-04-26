import { test, expect } from '../fixtures';

test.describe('Chat UI (fresh launch)', () => {
  test('renders header controls with no pre-seeded settings', async ({ taskpane }) => {
    await expect(taskpane.getByRole('link', { name: 'Copilot CLI plugin help' })).toBeVisible({
      timeout: 10_000,
    });
    await expect(taskpane.getByRole('button', { name: 'New conversation' })).toBeVisible({
      timeout: 10_000,
    });
  });

  test('shows the Composer input', async ({ taskpane }) => {
    await expect(taskpane.getByPlaceholder('Send a message...')).toBeVisible({ timeout: 10_000 });
  });

  test('shows the default agent picker', async ({ taskpane }) => {
    await expect(taskpane.getByRole('button', { name: 'Select agent' })).toBeVisible({
      timeout: 10_000,
    });
  });
});

test.describe('Chat UI (configured state)', () => {
  test('renders the chat header controls', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('link', { name: 'Copilot CLI plugin help' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'New conversation' })).toBeVisible();
  });

  test('Plugins button is NOT present — plugin management is done via CLI', async ({
    configuredTaskpane: page,
  }) => {
    // Plugin Hub was removed — users manage plugins via `copilot plugin add/update/remove`
    const pluginsButtons = page.getByRole('button', { name: 'Plugins' });
    await expect(pluginsButtons).toHaveCount(0);
  });

  test('MCP servers button is visible in the input toolbar', async ({
    configuredTaskpane: page,
  }) => {
    // The VS Code-style MCP server tools picker lives in the input toolbar,
    // using the server icon (not the extensions icon used by Plugins).
    await expect(page.getByRole('button', { name: 'MCP servers' })).toBeVisible();
  });

  test('MCP servers popover lists workiq and powerbi bundled servers', async ({
    configuredTaskpane: page,
  }) => {
    await page.getByRole('button', { name: 'MCP servers' }).click();
    // Both bundled MCP servers must appear in the popover
    await expect(page.getByRole('button', { name: 'workiq' })).toBeVisible({ timeout: 5000 });
    await expect(page.getByRole('button', { name: 'powerbi' })).toBeVisible({ timeout: 5000 });
  });

  test('does not show the removed skill picker button', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Agent skills' })).toHaveCount(0);
  });

  test('displays the model picker in the toolbar', async ({ configuredTaskpane: page }) => {
    // The model picker shows the active model name (default: Claude Sonnet 4)
    await expect(page.getByText('Claude Sonnet 4')).toBeVisible({ timeout: 5000 });
  });

  test('displays the agent picker', async ({ configuredTaskpane: page }) => {
    // The agent picker should show the active agent
    await expect(page.getByRole('button', { name: 'Select agent' })).toContainText('Default', {
      timeout: 5000,
    });
  });

  test('new conversation button is clickable', async ({ configuredTaskpane: page }) => {
    const btn = page.getByRole('button', { name: 'New conversation' });
    await expect(btn).toBeVisible();
    await btn.click();
    // No crash — composer input should still be functional
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible();
  });

  test('agent picker has no manage plugins button — plugin management is via CLI', async ({
    configuredTaskpane: page,
  }) => {
    await page.getByRole('button', { name: 'Select agent' }).click();
    await expect(page.getByRole('button', { name: 'Manage plugins…' })).toHaveCount(0);
    // Close the picker
    await page.keyboard.press('Escape');
  });

  test('/skills opens Copilot CLI skills management suggestions', async ({
    configuredTaskpane: page,
  }) => {
    const composer = page.getByRole('textbox', { name: 'Message input' });
    await composer.fill('/skills');
    await expect(page.getByRole('listbox', { name: 'skills command suggestions' })).toBeVisible();
    await expect(page.getByRole('option', { name: /\/skills list/i })).toBeVisible();
  });

  test('direct slash suggestions include installed Copilot CLI skills', async ({
    configuredTaskpane: page,
  }) => {
    const composer = page.getByRole('textbox', { name: 'Message input' });
    await composer.fill('/exc');
    await expect(page.getByRole('listbox', { name: 'slash suggestions' })).toBeVisible();
    await expect(page.getByRole('option').filter({ hasText: /\/excel/i }).first()).toBeVisible();
  });

  test('auto-scroll keeps thread pinned to newest content', async ({
    configuredTaskpane: page,
  }) => {
    // Build 60-message session history (30 pairs of user+assistant)
    const messages: unknown[] = [];
    for (let i = 0; i < 30; i++) {
      messages.push({
        id: `u-${i}`,
        role: 'user',
        content: [{ type: 'text', text: `User line ${i} ${'x'.repeat(60)}` }],
        createdAt: new Date(Date.now() - (60 - i) * 1000).toISOString(),
      });
      messages.push({
        id: `a-${i}`,
        role: 'assistant',
        content: [{ type: 'text', text: `Assistant line ${i} ${'y'.repeat(80)}` }],
        createdAt: new Date(Date.now() - (59 - i) * 1000).toISOString(),
      });
    }

    const historyJSON = JSON.stringify({
      state: {
        sessions: [
          {
            id: 'scroll-test-session',
            title: 'Scroll test',
            host: 'excel',
            updatedAt: Date.now(),
            messages,
          },
        ],
        activeSessionId: 'scroll-test-session',
      },
      version: 0,
    });

    // Register an init script that seeds session history into localStorage
    // BEFORE page code runs — addInitScript fires on every future navigation
    // including the reload below, so Zustand hydrates with the full session
    // on the very first render (no race condition).
    await page.addInitScript((json: string) => {
      localStorage.setItem('office-coding-agent-session-history', json);
    }, historyJSON);

    await page.reload();
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible();

    // Wait for at least one message element to be in the DOM, confirming that
    // session history was restored (not the empty welcome screen).
    await page.waitForSelector('[data-role="user"]', { timeout: 10_000 });

    // Wait (with retries) for the viewport to be scrolled to the bottom.
    // The MessageList auto-scroll effect should have fired by now.
    await page.waitForFunction(
      () => {
        const viewport = document.querySelector('.aui-thread-viewport') as HTMLElement | null;
        if (!viewport) return false;
        const delta = viewport.scrollHeight - viewport.scrollTop - viewport.clientHeight;
        return delta <= 8;
      },
      { timeout: 5_000 }
    );
  });
});
