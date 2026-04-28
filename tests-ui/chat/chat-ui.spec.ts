import { execFileSync } from 'node:child_process';
import { test, expect } from '../fixtures';

function getCliMcpServerNames(): string[] {
  const stdout = execFileSync('copilot', ['mcp', 'list', '--json'], {
    encoding: 'utf8',
    windowsHide: true,
  });
  const parsed = JSON.parse(stdout) as { mcpServers?: Record<string, unknown> };
  return Object.keys(parsed.mcpServers ?? {});
}

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

  test('shows the CLI-backed agent picker', async ({ taskpane }) => {
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

  test('MCP servers popover matches the Copilot CLI config', async ({ configuredTaskpane: page }) => {
    const cliServerNames = getCliMcpServerNames();

    await page.getByRole('button', { name: 'MCP servers' }).click();
    if (cliServerNames.length === 0) {
      await expect(page.getByText('No MCP servers available.')).toBeVisible({ timeout: 5000 });
      return;
    }

    for (const serverName of cliServerNames) {
      await expect(page.getByRole('button', { name: serverName })).toBeVisible({ timeout: 5000 });
    }
  });

  test('does not show the removed skill picker button', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Agent skills' })).toHaveCount(0);
  });

  test('displays the model picker in the toolbar', async ({ configuredTaskpane: page }) => {
    // The model picker shows the active model name (default: Claude Sonnet 4)
    await expect(page.getByText('Claude Sonnet 4')).toBeVisible({ timeout: 5000 });
  });

  test('displays the CLI-backed agent picker', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Select agent' })).toBeVisible();
    await page.getByRole('button', { name: 'Select agent' }).click();
    await expect(page.getByText('Agents')).toBeVisible();
  });

  test('new conversation button is clickable', async ({ configuredTaskpane: page }) => {
    const btn = page.getByRole('button', { name: 'New conversation' });
    await expect(btn).toBeVisible();
    await btn.click();
    // No crash — composer input should still be functional
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible();
  });

  test('does not show manage plugins button — plugin management is via CLI', async ({
    configuredTaskpane: page,
  }) => {
    await expect(page.getByRole('button', { name: 'Manage plugins…' })).toHaveCount(0);
  });

  test('/skills does not expose management command suggestions', async ({
    configuredTaskpane: page,
  }) => {
    const composer = page.getByRole('textbox', { name: 'Message input' });
    await composer.fill('/skills');
    await expect(page.getByRole('listbox', { name: 'slash suggestions' })).toBeVisible();
    await expect(page.getByRole('option', { name: /\/skills list/i })).toHaveCount(0);
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
