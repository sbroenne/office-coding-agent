/**
 * Playwright tests for the AssistantMessage action bar.
 *
 * Verifies:
 *  - Copy / thumbsup / thumbsdown buttons appear when hovering over a
 *    completed assistant message.
 *  - Clicking the Copy button places the message text on the clipboard.
 *  - No crash (error boundary) appears after a tool-call-only response that
 *    has no trailing text part.
 *
 * Uses the REAL Copilot API through the dev server (no mocks).
 * Requires `npm run dev` to be running on https://localhost:3000.
 */

import { test, expect } from '../fixtures';

const AI_TIMEOUT = 45_000;

/**
 * Helper: send a prompt, wait for the response to complete, and return the
 * last assistant message locator.
 */
async function sendAndWaitForResponse(page: import('@playwright/test').Page, prompt: string) {
  const composer = page.getByPlaceholder('Send a message...');
  await expect(composer).toBeVisible({ timeout: 10_000 });
  await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
  await expect(page.getByText('Connection failed')).not.toBeVisible();

  await composer.fill(prompt);
  await composer.press('Enter');

  // Wait for the Stop/Cancel button to disappear — response is complete
  await expect(page.getByRole('button', { name: /^(Stop|Cancel)$/ })).not.toBeVisible({
    timeout: AI_TIMEOUT,
  });

  const assistantMsg = page.locator('[data-role="assistant"]').last();
  await expect(assistantMsg).toBeVisible({ timeout: 5_000 });
  return assistantMsg;
}

test.describe('Action bar (live Copilot)', () => {
  test('copy + thumbsup + thumbsdown buttons appear when hovering over a completed response', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const assistantMsg = await sendAndWaitForResponse(page, 'Reply with exactly one word: READY');

    // Hover over the assistant message to reveal the action bar
    await assistantMsg.hover();

    const actionBar = assistantMsg.locator('.aui-assistant-action-bar');
    await expect(actionBar).toBeVisible({ timeout: 3_000 });

    // All three action buttons must be visible
    await expect(page.getByRole('button', { name: 'Copy' })).toBeVisible({ timeout: 3_000 });
    await expect(page.getByRole('button', { name: 'Good response' })).toBeVisible({
      timeout: 3_000,
    });
    await expect(page.getByRole('button', { name: 'Bad response' })).toBeVisible({
      timeout: 3_000,
    });
  });

  test('action bar is hidden before hover (opacity-0 class set)', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const assistantMsg = await sendAndWaitForResponse(page, 'Reply with exactly one word: CHECK');

    const actionBar = assistantMsg.locator('.aui-assistant-action-bar');

    // Without hovering, the action bar should be in the DOM but invisible
    await expect(actionBar).toBeAttached({ timeout: 3_000 });
    await expect(actionBar).toHaveClass(/opacity-0/);
  });

  test('clicking Copy puts the response text on the clipboard', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    // Grant clipboard-read permission for this context
    await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);

    const assistantMsg = await sendAndWaitForResponse(
      page,
      'Reply with exactly one word: CLIPBOARD'
    );

    // Hover to reveal the action bar, then click Copy
    await assistantMsg.hover();
    await expect(page.getByRole('button', { name: 'Copy' })).toBeVisible({ timeout: 3_000 });
    await page.getByRole('button', { name: 'Copy' }).click();

    // Clipboard should contain a non-empty string from the response
    const clipText = await page.evaluate(() => navigator.clipboard.readText());
    expect(clipText.trim().length).toBeGreaterThan(0);
  });

  test('no error boundary appears after a tool-call-only response (crash regression)', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    // This prompt triggers manage_skills (a pure tool call).
    // Some Copilot responses end with the tool result and no trailing text
    // part — this was the scenario that caused the "MessagePartText can only
    // be used inside text or reasoning message parts" crash.
    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 10_000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });

    await composer.fill(
      'Use the manage_skills tool with action "list". Only call the tool, do not add any text response after it.'
    );
    await composer.press('Enter');

    // Wait for response to complete
    await expect(page.getByRole('button', { name: /^(Stop|Cancel)$/ })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // Critical: error boundary must not have triggered
    await expect(page.getByText(/something went wrong/i)).not.toBeVisible({ timeout: 2_000 });
    await expect(page.getByText(/MessagePartText can only/i)).not.toBeVisible({ timeout: 2_000 });

    // The thread is still functional — composer is visible
    await expect(page.getByPlaceholder('Send a message...')).toBeVisible({ timeout: 3_000 });
  });

  test('action bar is not shown while a response is streaming (hideWhenRunning)', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 10_000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });

    await composer.fill('Count from 1 to 5, one number per line.');
    await composer.press('Enter');

    // While streaming: the Stop/Cancel button is visible
    const stopBtn = page.getByRole('button', { name: /^(Stop|Cancel)$/ });
    await expect(stopBtn).toBeVisible({ timeout: 10_000 });

    // While running, hover over the in-progress assistant message
    const assistantMsg = page.locator('[data-role="assistant"]').last();
    if (await assistantMsg.isVisible()) {
      await assistantMsg.hover();
      // Action bar should NOT be visible during streaming (hideWhenRunning)
      const actionBar = assistantMsg.locator('.aui-assistant-action-bar');
      if (await actionBar.isAttached()) {
        // If attached, it must be hidden (data-state or display), not visible buttons
        await expect(page.getByRole('button', { name: 'Copy' })).not.toBeVisible();
      }
    }

    // Wait for streaming to finish, then re-verify action bar appears post-hover
    await expect(stopBtn).not.toBeVisible({ timeout: AI_TIMEOUT });
    await assistantMsg.hover();
    await expect(page.getByRole('button', { name: 'Copy' })).toBeVisible({ timeout: 3_000 });
  });
});
