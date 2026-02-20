import { test, expect } from '../fixtures';

test.describe('Settings Dialog', () => {
  async function openSettings(page: import('@playwright/test').Page) {
    const header = page.locator('div.flex.items-center.justify-between');
    await header.locator('button').last().click();
    await expect(page.getByRole('heading', { name: 'Settings' })).toBeVisible({ timeout: 3000 });
  }

  test('opens and shows the configured endpoint', async ({ configuredTaskpane: page }) => {
    await openSettings(page);
    await expect(page.getByRole('heading', { name: 'Test Endpoint' })).toBeVisible();
  });

  test('shows endpoint URL in the dialog', async ({ configuredTaskpane: page }) => {
    await openSettings(page);
    await expect(page.getByText('test.openai.azure.com')).toBeVisible();
  });
});
