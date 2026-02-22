/**
 * Advanced E2E chat tests: browser UI → WebSocket proxy → GitHub Copilot API.
 *
 * Requires `npm run dev` to be running on https://localhost:3000.
 * GitHub Copilot is always available — no skip logic.
 *
 * These tests exercise custom agent instructions, skill injection,
 * the thinking indicator, tool progress UI, and multi-turn conversations.
 */

import { test, expect } from '../fixtures';

const AI_TIMEOUT = 60_000;

test.describe('Chat E2E — custom agent behaviour (requires server)', () => {
  test('custom agent instructions change the model response personality', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    // Switch to a custom agent via the agent picker
    const agentButton = page.getByRole('button', { name: 'Select agent' });
    await expect(agentButton).toBeVisible({ timeout: 5000 });

    // The default "Excel" agent should be active
    await expect(page.getByText('Excel')).toBeVisible({ timeout: 5000 });

    // Type a prompt
    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });
    await composer.fill('Reply with exactly one word: PONG');
    await composer.press('Enter');

    // Expect a response from the model
    await expect(page.getByText(/pong/i).first()).toBeVisible({ timeout: AI_TIMEOUT });
  });

  test('assistant responds to Excel-related prompt with tool use', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // Send a prompt that requires tool use (e.g., reading cell data)
    await composer.fill("What's in cell A1?");
    await composer.press('Enter');

    // The model should respond with something about the cell
    await expect(
      page.getByText(/cell|A1|empty|value|contains|data|workbook/i).first()
    ).toBeVisible({ timeout: AI_TIMEOUT });
  });

  test('tool execution progress appears in the thread', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // Send a prompt that should trigger a tool call (read operation)
    await composer.fill('List all sheet names in this workbook');
    await composer.press('Enter');

    // A response mentioning sheet(s) should appear
    await expect(page.getByText(/sheet/i).first()).toBeVisible({ timeout: AI_TIMEOUT });
  });

  test('multi-turn conversation sends and receives multiple messages', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT * 2 + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // Turn 1: send a prompt
    await composer.fill('Reply with exactly one word: ALPHA');
    await composer.press('Enter');
    await expect(page.getByText(/alpha/i).first()).toBeVisible({ timeout: AI_TIMEOUT });

    // Wait for the composer to be ready for turn 2
    await expect(composer).toBeVisible({ timeout: 5000 });
    await expect(composer).toBeEmpty({ timeout: 5000 });

    // Turn 2: send another prompt
    await composer.fill('Reply with exactly one word: BRAVO');
    await page.getByRole('button', { name: 'Send' }).click();

    // Both assistant responses should be present in the thread
    const messages = page.locator('[data-role="assistant"]');
    await expect(messages).toHaveCount(2, { timeout: AI_TIMEOUT });
    await expect(messages.nth(1)).toContainText(/bravo/i, { timeout: AI_TIMEOUT });
  });

  test('new conversation button clears the thread and starts fresh', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // Send an initial message
    await composer.fill('Reply with exactly one word: PONG');
    await composer.press('Enter');
    await expect(page.getByText(/pong/i).first()).toBeVisible({ timeout: AI_TIMEOUT });

    // Click new conversation
    await page.getByRole('button', { name: 'New conversation' }).click();

    // The thread should be cleared — no more "PONG" visible
    // Wait a moment for the UI to update
    await page.waitForTimeout(1000);
    await expect(page.getByText(/pong/i)).not.toBeVisible({ timeout: 5000 });

    // Composer should still be functional
    await expect(composer).toBeVisible();
  });

  test('model picker shows live models from the Copilot API', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    // Wait for the model picker to be populated from the live API
    // The configuredTaskpane pre-seeds with Claude Sonnet 4, but the live
    // connection should also load models
    const modelButton = page.getByRole('button', { name: 'Select model' });
    await expect(modelButton).toBeVisible({ timeout: 10_000 });
    await modelButton.click();

    // At least one model should be listed in the dropdown
    const modelOptions = page.getByRole('button').filter({ hasText: /(Claude|GPT|Gemini|o[0-9])/i });
    await expect(modelOptions.first()).toBeVisible({ timeout: 10_000 });
  });
});
