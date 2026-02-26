import { test, expect } from '../fixtures';

test.describe('Chat UI (fresh launch)', () => {
  test('renders header controls with no pre-seeded settings', async ({ taskpane }) => {
    await expect(taskpane.getByRole('button', { name: 'Agent skills' })).toBeVisible({
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
    await expect(page.getByRole('button', { name: 'Agent skills' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'New conversation' })).toBeVisible();
  });

  test('shows the skill picker button', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Agent skills' })).toBeVisible();
  });

  test('displays the model picker in the toolbar', async ({ configuredTaskpane: page }) => {
    // The model picker shows the active model name (default: Claude Sonnet 4)
    await expect(page.getByText('Claude Sonnet 4')).toBeVisible({ timeout: 5000 });
  });

  test('displays the agent picker', async ({ configuredTaskpane: page }) => {
    // The agent picker should show the active agent
    await expect(page.getByText('Excel')).toBeVisible({ timeout: 5000 });
  });

  test('new conversation button is clickable', async ({ configuredTaskpane: page }) => {
    const btn = page.getByRole('button', { name: 'New conversation' });
    await expect(btn).toBeVisible();
    await btn.click();
    // No crash — composer input should still be functional
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible();
  });

  test('agent manager dialog supports keyboard open/close', async ({
    configuredTaskpane: page,
  }) => {
    await page.getByRole('button', { name: 'Select agent' }).click();

    const manageAgents = page.getByRole('button', { name: 'Manage agents…' });
    await manageAgents.focus();
    await page.keyboard.press('Enter');

    await expect(page.getByRole('heading', { name: 'Manage Agents' })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('heading', { name: 'Manage Agents' })).not.toBeVisible();
  });

  test('skill manager dialog supports keyboard open/close', async ({
    configuredTaskpane: page,
  }) => {
    await page.getByRole('button', { name: 'Agent skills' }).click();

    const manageSkills = page.getByRole('button', { name: 'Manage skills…' });
    await manageSkills.focus();
    await page.keyboard.press('Enter');

    await expect(page.getByRole('heading', { name: 'Manage Skills' })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('heading', { name: 'Manage Skills' })).not.toBeVisible();
  });

  test('auto-scroll keeps thread pinned to newest content', async ({
    configuredTaskpane: page,
  }) => {
    await page.evaluate(() => {
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

      localStorage.setItem(
        'office-coding-agent-session-history',
        JSON.stringify({
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
        })
      );
    });

    await page.reload();
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible();

    const pinnedToBottom = await page.evaluate(async () => {
      await new Promise(resolve => requestAnimationFrame(() => resolve(undefined)));
      const viewport = document.querySelector('.aui-thread-viewport') as HTMLElement | null;
      if (!viewport) return false;
      const delta = viewport.scrollHeight - viewport.scrollTop - viewport.clientHeight;
      return delta <= 8;
    });

    expect(pinnedToBottom).toBe(true);
  });
});
