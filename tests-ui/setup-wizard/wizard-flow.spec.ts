import { test, expect } from '../fixtures';

test.describe('Setup Wizard', () => {
  test('shows the endpoint step on first launch', async ({ taskpane }) => {
    await expect(taskpane.getByText('Connect to Azure AI Foundry')).toBeVisible();
    await expect(taskpane.getByLabel('Resource URL')).toBeVisible();
  });

  test('Next button enables when URL is entered', async ({ taskpane }) => {
    // Clear any build-time default, then type a URL
    await taskpane.getByLabel('Resource URL').clear();
    await expect(taskpane.getByRole('button', { name: 'Next' })).toBeDisabled();

    await taskpane.getByLabel('Resource URL').fill('https://my-resource.openai.azure.com');
    await expect(taskpane.getByRole('button', { name: 'Next' })).toBeEnabled();
  });

  test('advances to the auth step', async ({ taskpane }) => {
    await taskpane.getByLabel('Resource URL').fill('https://my-resource.openai.azure.com');
    await taskpane.getByRole('button', { name: 'Next' }).click();

    await expect(taskpane.getByText('Authentication')).toBeVisible();
    await expect(taskpane.getByRole('textbox', { name: 'API Key' })).toBeVisible();
    await expect(taskpane.getByRole('button', { name: 'Connect' })).toBeDisabled();
  });

  test('Connect button enables when API key is entered', async ({ taskpane }) => {
    await taskpane.getByLabel('Resource URL').fill('https://my-resource.openai.azure.com');
    await taskpane.getByRole('button', { name: 'Next' }).click();

    await taskpane.getByRole('textbox', { name: 'API Key' }).fill('test-api-key-12345');
    await expect(taskpane.getByRole('button', { name: 'Connect' })).toBeEnabled();
  });

  test('can toggle API key visibility', async ({ taskpane }) => {
    await taskpane.getByLabel('Resource URL').fill('https://my-resource.openai.azure.com');
    await taskpane.getByRole('button', { name: 'Next' }).click();

    const apiKeyInput = taskpane.getByRole('textbox', { name: 'API Key' });
    await expect(apiKeyInput).toHaveAttribute('type', 'password');

    await taskpane.getByRole('button', { name: 'Show API key' }).click();
    await expect(apiKeyInput).toHaveAttribute('type', 'text');

    await taskpane.getByRole('button', { name: 'Hide API key' }).click();
    await expect(apiKeyInput).toHaveAttribute('type', 'password');
  });

  test('Back button returns to endpoint step', async ({ taskpane }) => {
    await taskpane.getByLabel('Resource URL').fill('https://my-resource.openai.azure.com');
    await taskpane.getByRole('button', { name: 'Next' }).click();
    await expect(taskpane.getByText('Authentication')).toBeVisible();

    await taskpane.getByRole('button', { name: 'Back' }).click();
    await expect(taskpane.getByText('Connect to Azure AI Foundry')).toBeVisible();
    // URL should still be filled
    await expect(taskpane.getByLabel('Resource URL')).toHaveValue(
      'https://my-resource.openai.azure.com'
    );
  });

  test('shows connecting state after clicking Connect', async ({ taskpane }) => {
    await taskpane.getByLabel('Resource URL').fill('https://my-resource.openai.azure.com');
    await taskpane.getByRole('button', { name: 'Next' }).click();

    await taskpane.getByRole('textbox', { name: 'API Key' }).fill('test-api-key-12345');
    await taskpane.getByRole('button', { name: 'Connect' }).click();

    // Depending on response speed, we may briefly show "Connecting..." or move
    // directly to the next wizard step.
    const transitionState = taskpane
      .getByText('Connecting...')
      .or(taskpane.getByText('Add Your Models'))
      .or(taskpane.getByText("You're all set!"));
    await expect(transitionState).toBeVisible();
  });
});
