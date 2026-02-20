import { test, expect } from '../fixtures';

/**
 * Tests for the app in "fresh launch" state (no pre-seeded settings).
 * With the Copilot SDK migration the SetupWizard was removed — the app
 * goes straight to the chat UI using GitHub Copilot CLI for authentication.
 */
test.describe('App fresh launch', () => {
  test('loads and shows the chat header controls', async ({ taskpane }) => {
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

  test('shows the default model picker', async ({ taskpane }) => {
    // Default model is claude-sonnet-4.5 → displayed as "Claude Sonnet 4.5"
    await expect(taskpane.getByText('Claude Sonnet 4.5')).toBeVisible({ timeout: 10_000 });
  });

  test('shows the default agent picker', async ({ taskpane }) => {
    await expect(taskpane.getByRole('button', { name: 'Select agent' })).toBeVisible({
      timeout: 10_000,
    });
  });

  test('shows the New conversation button', async ({ taskpane }) => {
    await expect(taskpane.getByRole('button', { name: 'New conversation' })).toBeVisible({
      timeout: 10_000,
    });
  });
});
