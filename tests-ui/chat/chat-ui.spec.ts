import { test, expect } from '../fixtures';

test.describe('Chat UI (configured state)', () => {
  test('renders the chat header controls', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Agent skills' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'New conversation' })).toBeVisible();
  });

  test('shows the skill picker button', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Agent skills' })).toBeVisible();
  });

  test('displays the model picker in the toolbar', async ({ configuredTaskpane: page }) => {
    // The model picker shows the active model name (default: Claude Sonnet 4.5)
    await expect(page.getByText('Claude Sonnet 4.5')).toBeVisible({ timeout: 5000 });
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

  test('agent manager dialog supports keyboard open/close', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Select agent' }).click();

    const manageAgents = page.getByRole('button', { name: 'Manage agents…' });
    await manageAgents.focus();
    await page.keyboard.press('Enter');

    await expect(page.getByRole('heading', { name: 'Manage Agents' })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('heading', { name: 'Manage Agents' })).not.toBeVisible();
  });

  test('skill manager dialog supports keyboard open/close', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Agent skills' }).click();

    const manageSkills = page.getByRole('button', { name: 'Manage skills…' });
    await manageSkills.focus();
    await page.keyboard.press('Enter');

    await expect(page.getByRole('heading', { name: 'Manage Skills' })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('heading', { name: 'Manage Skills' })).not.toBeVisible();
  });
});
