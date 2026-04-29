import { test, expect } from '../fixtures';

test.describe('Permissions and session history UI', () => {
  test.describe.configure({ timeout: 120_000 });

  test('opens permissions panel and can manage SDK approvals', async ({
    configuredTaskpane: page,
  }) => {
    await page.getByRole('button', { name: 'Permissions' }).click();

    await expect(page.getByRole('heading', { name: 'Permissions' })).toBeVisible();
    await expect(
      page.getByText('Manage Copilot CLI approvals')
    ).toBeVisible();

    await expect(page.getByText('Allow all')).toBeVisible();
    await expect(page.getByText('Session approvals', { exact: true })).toBeVisible();

    await page.getByRole('button', { name: 'Off' }).click();
    await expect(page.getByRole('button', { name: 'On' })).toBeVisible();

    await page.getByRole('button', { name: 'Reset session approvals' }).click();
  });

  test('opens history manage panel and deletes a saved session', async ({
    configuredTaskpane: page,
  }) => {
    await page.evaluate(() => {
      const payload = {
        version: 0,
        state: {
          sessions: [
            {
              id: 's-excel-1',
              title: 'Budget review session',
              host: 'excel',
              updatedAt: Date.now() - 120000,
              messages: [],
            },
            {
              id: 's-excel-2',
              title: 'Pipeline follow-up',
              host: 'excel',
              updatedAt: Date.now() - 60000,
              messages: [],
            },
            {
              id: 's-word-1',
              title: 'Word-only session',
              host: 'word',
              updatedAt: Date.now() - 30000,
              messages: [],
            },
          ],
          activeSessionId: 's-excel-2',
        },
      };
      localStorage.setItem('office-coding-agent-session-history', JSON.stringify(payload));
      const runtime = (
        globalThis as {
          OfficeRuntime?: { storage?: { setItem?: (k: string, v: string) => Promise<void> } };
        }
      ).OfficeRuntime;
      if (runtime?.storage?.setItem) {
        void runtime.storage.setItem(
          'office-coding-agent-session-history',
          JSON.stringify(payload)
        );
      }
    });

    await page.reload();
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible();

    await page.getByRole('button', { name: 'Session history' }).click();
    await expect(page.getByRole('button', { name: 'Manage history…' })).toBeVisible();

    await page.getByRole('button', { name: 'Manage history…' }).click();
    await expect(page.getByRole('heading', { name: 'Session History' })).toBeVisible();
    await expect(page.getByText('Budget review session')).toBeVisible();

    await expect(page.getByText('Word-only session')).not.toBeVisible();

    await page.getByRole('button', { name: 'Delete session' }).nth(1).click();

    await expect(page.getByText('Budget review session')).not.toBeVisible();
  });
});
