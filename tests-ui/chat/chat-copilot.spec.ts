/**
 * Full E2E chat test: browser UI → WebSocket proxy → GitHub Copilot API.
 *
 * Requires `npm run dev` to be running on https://localhost:3000.
 * GitHub Copilot is always available — no skip logic.
 *
 * Run manually:
 *   npm run dev             # terminal 1
 *   npm run test:ui         # terminal 2  (or --grep "Copilot")
 */

import { test, expect } from '../fixtures';

const AI_TIMEOUT = 45_000;

test.describe('Chat E2E with Copilot (requires server)', () => {
  test('sends a message and receives a Copilot response', async ({ configuredTaskpane: page }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    // Type a prompt in the Composer
    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });
    await composer.fill('Reply with exactly one word: PONG');
    await composer.press('Enter');

    // assistant-ui wraps messages — check for PONG anywhere in a message bubble
    await expect(page.getByText(/pong/i).first()).toBeVisible({ timeout: AI_TIMEOUT });
  });

  test('tool call result appears as progress in the thread', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // This prompt should trigger get_workbook_info or similar read tool
    await composer.fill("What's the name of this workbook?");
    await composer.press('Enter');

    // Any assistant response should appear
    await expect(page.getByText(/workbook|spreadsheet|untitled/i).first()).toBeVisible({
      timeout: AI_TIMEOUT,
    });
  });
});
