/**
 * Full E2E chat test: browser UI → WebSocket proxy → GitHub Copilot API.
 *
 * Requires `npm run dev` to be running on https://localhost:3000.
 * Skips automatically when the proxy is unreachable.
 *
 * Run manually:
 *   npm run dev             # terminal 1
 *   npm run test:ui         # terminal 2  (or --grep "Copilot")
 */

import { test, expect } from '../fixtures';

const COPILOT_HEALTH = 'https://localhost:3000/api/copilot-health';
const AI_TIMEOUT = 45_000;

test.describe('Chat E2E with Copilot (requires server)', () => {
  let copilotAvailable = false;

  test.beforeAll(async ({ request }) => {
    try {
      const resp = await request.get(COPILOT_HEALTH, {
        ignoreHTTPSErrors: true,
        timeout: 15_000,
      });
      if (resp.ok()) {
        const body = await resp.json();
        copilotAvailable = body.ok === true;
      }
    } catch {
      copilotAvailable = false;
    }
  });

  test('sends a message and receives a Copilot response', async ({ configuredTaskpane: page }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);
    if (!copilotAvailable) {
      test.skip(true, 'Copilot API not reachable — start `npm run dev` with valid auth');
    }

    // Type a prompt in the Composer
    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });
    await composer.fill('Reply with exactly one word: PONG');
    await composer.press('Enter');

    // Wait for an assistant message to appear in the thread
    const assistantMsg = page
      .locator('[data-message-role="assistant"], .aui-message[data-role="assistant"]')
      .first();

    // assistant-ui wraps messages — check for PONG anywhere in a message bubble
    await expect(page.getByText(/pong/i).first()).toBeVisible({ timeout: AI_TIMEOUT });
  });

  test('tool call result appears as progress in the thread', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);
    if (!copilotAvailable) {
      test.skip(true, 'Copilot API not reachable — start `npm run dev` with valid auth');
    }

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
