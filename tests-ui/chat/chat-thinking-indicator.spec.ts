/**
 * Tests for the thinking indicator UI lifecycle.
 *
 * Uses the REAL Copilot API through the dev server (no mocks).
 * Requires `npm run dev` to be running on https://localhost:3000.
 */

import { test, expect } from '../fixtures';

const AI_TIMEOUT = 45_000;

test.describe('Thinking indicator (live Copilot)', () => {
  test('thinking indicator shows dynamic text during tool execution', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // Wait for the WebSocket session to be established before sending.
    // The app shows "Connecting to Copilot..." during handshake; we wait for
    // it to disappear so that sessionRef.current is populated.
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    // Capture all thinking indicator text values via MutationObserver
    await page.evaluate(() => {
      (window as unknown as Record<string, string[]>).__thinkingTexts = [];
      const observer = new MutationObserver(() => {
        const el = document.querySelector('.aui-assistant-thinking-indicator span');
        if (el?.textContent) {
          const texts = (window as unknown as Record<string, string[]>).__thinkingTexts;
          const last = texts[texts.length - 1];
          if (el.textContent !== last) {
            texts.push(el.textContent);
          }
        }
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    });

    // Prompt that triggers manage_skills tool (no Excel needed)
    await composer.fill('Use the manage_skills tool with action "list" and tell me how many skills you have.');
    await composer.press('Enter');

    // Wait for the response to complete
    await expect(page.getByRole('button', { name: 'Cancel' })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // Retrieve captured thinking texts
    const thinkingTexts = await page.evaluate(
      () => (window as unknown as Record<string, string[]>).__thinkingTexts
    );
    console.log('  Captured thinking texts:', thinkingTexts);

    // Should capture at least one thinking indicator text while running
    expect(thinkingTexts.length).toBeGreaterThanOrEqual(1);
  });

  test('report_intent does NOT create a tool-call card in the thread', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });

    // Wait for the WebSocket session to be established
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    // Use manage_skills — a tool that doesn't need Excel
    await composer.fill('Use the manage_skills tool with action "list" and tell me how many skills you have.');
    await composer.press('Enter');

    // Wait for the response to complete — Cancel button disappears
    await expect(page.getByRole('button', { name: 'Cancel' })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // Verify the assistant responded
    const assistantMsg = page.locator('[data-role="assistant"]');
    await expect(assistantMsg.first()).toBeVisible({ timeout: 5_000 });

    // No tool card should mention "report_intent" — it's an internal event
    const toolCards = page.locator('[data-slot="tool-fallback-root"]');
    const count = await toolCards.count();
    for (let i = 0; i < count; i++) {
      await expect(toolCards.nth(i)).not.toContainText(/report.intent/i);
    }
  });

  test('thinking indicator renders inline in message stream, not inside the sticky footer', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    await composer.fill('Use the manage_skills tool with action "list" and tell me how many skills you have.');
    await composer.press('Enter');

    // Wait for the thinking indicator to appear
    const indicator = page.locator('.aui-assistant-thinking-indicator');
    await expect(indicator).toBeVisible({ timeout: AI_TIMEOUT });

    // The indicator must NOT be a descendant of the viewport footer
    const isInsideFooter = await indicator.evaluate(el =>
      !!el.closest('.aui-thread-viewport-footer')
    );
    expect(isInsideFooter).toBe(false);

    // The indicator must be a descendant of the scrollable viewport
    const isInsideViewport = await indicator.evaluate(el =>
      !!el.closest('.aui-thread-viewport')
    );
    expect(isInsideViewport).toBe(true);

    // The indicator must be rendered within an assistant message block
    const isInsideAssistantMessage = await indicator.evaluate(el =>
      !!el.closest('.aui-assistant-message-root')
    );
    expect(isInsideAssistantMessage).toBe(true);

    // Geometric guard: indicator must render above the composer area
    const isAboveComposer = await page.evaluate(() => {
      const indicatorEl = document.querySelector('.aui-assistant-thinking-indicator');
      const composerEl = document.querySelector('.aui-composer-root');
      if (!indicatorEl || !composerEl) return false;
      const indicatorRect = indicatorEl.getBoundingClientRect();
      const composerRect = composerEl.getBoundingClientRect();
      return indicatorRect.bottom <= composerRect.top;
    });
    expect(isAboveComposer).toBe(true);

    // The indicator must appear BELOW the last message in DOM order
    const isAfterMessages = await page.evaluate(() => {
      const messages = document.querySelector('[data-role="user"]');
      const ind = document.querySelector('.aui-assistant-thinking-indicator');
      if (!messages || !ind) return false;
      return !!(messages.compareDocumentPosition(ind) & Node.DOCUMENT_POSITION_FOLLOWING);
    });
    expect(isAfterMessages).toBe(true);

    // Wait for response to finish
    await expect(page.getByRole('button', { name: 'Cancel' })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // After completion, indicator must be gone
    await expect(indicator).not.toBeVisible();
  });
});
