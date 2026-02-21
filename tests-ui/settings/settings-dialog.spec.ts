import { test, expect } from '../fixtures';

test.describe('Model Picker', () => {
  test('shows the active model name', async ({ configuredTaskpane: page }) => {
    await expect(page.getByText('Claude Sonnet 4')).toBeVisible({ timeout: 5000 });
  });

  test('opens the model list on click', async ({ configuredTaskpane: page }) => {
    // Click the model picker button (contains the model name text)
    await page.getByText('Claude Sonnet 4').click();

    // Dropdown should show model groups
    await expect(page.getByText('Anthropic')).toBeVisible({ timeout: 3000 });
    await expect(page.getByText('OpenAI')).toBeVisible({ timeout: 3000 });
  });

  test('can select a different model', async ({ configuredTaskpane: page }) => {
    await page.getByText('Claude Sonnet 4').click();
    await page.getByText('GPT-4.1').click();

    // Picker now shows the newly selected model
    await expect(page.getByText('GPT-4.1')).toBeVisible({ timeout: 3000 });
  });
});
