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
        const el = document.querySelector('.aui-thinking-indicator span');
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

    // Should have at least one non-"Thinking…" text (tool name or intent)
    const dynamicTexts = thinkingTexts.filter(t => t !== 'Thinking…');
    expect(dynamicTexts.length).toBeGreaterThanOrEqual(1);
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
});
