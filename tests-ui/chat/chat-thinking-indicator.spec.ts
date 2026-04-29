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

    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    // Capture all thinking indicator text values via MutationObserver
    // Now captures both the inline shimmer AND the Working box spinner
    await page.evaluate(() => {
      (window as unknown as Record<string, string[]>).__thinkingTexts = [];
      const observer = new MutationObserver(() => {
        const texts = (window as unknown as Record<string, string[]>).__thinkingTexts;
        // Check inline shimmer
        const el = document.querySelector('.inline-working-progress .progress-step');
        if (el?.textContent) {
          const last = texts[texts.length - 1];
          if (el.textContent !== last) {
            texts.push(el.textContent);
          }
        }
        // Check Working box spinner
        const spinner = document.querySelector('[data-testid="working-spinner"] .chat-thinking-spinner-label');
        if (spinner?.textContent) {
          const last = texts[texts.length - 1];
          if (spinner.textContent !== last) {
            texts.push(spinner.textContent);
          }
        }
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    });

    // Prompt that triggers manage_memory tool (no Excel needed)
    await composer.fill('In one short paragraph, explain how Excel recalculates formulas.');
    await composer.press('Enter');

    // Wait for the response to complete
    await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({
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

    // Use manage_memory — a tool that doesn't need Excel
    await composer.fill(
      'Use the manage_memory tool with action "list" and tell me how many skills you have.'
    );
    await composer.press('Enter');

    // Wait for the response to complete — Cancel button disappears
    await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({
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

  test('thinking indicator renders inline in assistant message, not inside the sticky footer', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    await composer.fill(
      'Use the manage_memory tool with action "list" and tell me how many skills you have.'
    );
    await composer.press('Enter');

    // Wait for either the inline shimmer OR the Working box to appear
    const indicator = page.locator('.inline-working-progress');
    const workingBox = page.locator('.chat-thinking-box');
    await page.waitForFunction(
      () =>
        Boolean(
          document.querySelector('.inline-working-progress') ||
            document.querySelector('.chat-thinking-box')
        ),
      undefined,
      { timeout: AI_TIMEOUT }
    );

    // Whichever appears must be inside the assistant message (not in the footer)
    const progressElement = (await indicator.isVisible()) ? indicator : workingBox;
    const isInsideFooter = await progressElement.evaluate(
      el => !!el.closest('.aui-thread-viewport-footer')
    );
    expect(isInsideFooter).toBe(false);

    // The indicator must be a descendant of the scrollable viewport
    const isInsideViewport = await progressElement.evaluate(
      el => !!el.closest('.aui-thread-viewport')
    );
    expect(isInsideViewport).toBe(true);

    // The indicator must be rendered within an assistant message block
    const isInsideAssistantMessage = await progressElement.evaluate(
      el => !!el.closest('.aui-assistant-message-root')
    );
    expect(isInsideAssistantMessage).toBe(true);

    // Geometric guard: indicator must render above the composer area
    const isAboveComposer = await page.evaluate(() => {
      const indicatorEl =
        document.querySelector('.inline-working-progress') ??
        document.querySelector('.chat-thinking-box');
      const composerEl = document.querySelector('.aui-composer-root');
      if (!indicatorEl || !composerEl) return false;
      const indicatorRect = indicatorEl.getBoundingClientRect();
      const composerRect = composerEl.getBoundingClientRect();
      return indicatorRect.bottom <= composerRect.top;
    });
    expect(isAboveComposer).toBe(true);

    // The indicator must appear BELOW the last user message in DOM order
    const isAfterMessages = await page.evaluate(() => {
      const messages = document.querySelector('[data-role="user"]');
      const ind =
        document.querySelector('.inline-working-progress') ??
        document.querySelector('.chat-thinking-box');
      if (!messages || !ind) return false;
      return !!(messages.compareDocumentPosition(ind) & Node.DOCUMENT_POSITION_FOLLOWING);
    });
    expect(isAfterMessages).toBe(true);

    // Wait for response to finish
    await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // After completion, shimmer indicator must be gone
    await expect(indicator).not.toBeVisible();
  });

  test('Working box shows spinner between tool completions and collapses when done', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 5000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    // Capture Working box state transitions via MutationObserver
    await page.evaluate(() => {
      (window as unknown as Record<string, string[]>).__workingStates = [];
      const observer = new MutationObserver(() => {
        const spinner = document.querySelector('[data-testid="working-spinner"]');
        const workingTitle = document.querySelector('.chat-thinking-title-shimmer');
        const doneTitle = document.querySelector('.chat-thinking-title-done');
        const states = (window as unknown as Record<string, string[]>).__workingStates;
        if (spinner) {
          const text = spinner.textContent || '';
          const last = states[states.length - 1];
          if (text !== last) states.push(`spinner: ${text}`);
        }
        if (workingTitle) {
          const last = states[states.length - 1];
          if (last !== 'working-active') states.push('working-active');
        }
        if (doneTitle) {
          const text = doneTitle.textContent || '';
          const last = states[states.length - 1];
          if (`done: ${text}` !== last) states.push(`done: ${text}`);
        }
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    });

    // Prompt that triggers tool calls
    await composer.fill(
      'Use the manage_memory tool with action "list" and then use manage_memory with action "search" and query "missing". Tell me the counts.'
    );
    await composer.press('Enter');

    // Wait for response to complete
    await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // After completion, verify we captured Working box transitions
    const workingStates = await page.evaluate(
      () => (window as unknown as Record<string, string[]>).__workingStates
    );
    console.log('  Working box states:', workingStates);

    // Should have seen at least: working-active and done state
    expect(workingStates.length).toBeGreaterThanOrEqual(1);

    // After completion: Working box title should be visible (done state, not shimmer)
    const doneTitle = page.locator('.chat-thinking-title-done');
    await expect(doneTitle).toBeVisible({ timeout: 5000 });
    const doneText = await doneTitle.textContent();
    // Done label is the phase label or fallback "Working"
    expect(doneText).toBeTruthy();

    // Take a visual regression screenshot of the completed Working box
    const assistantMessage = page.locator('[data-role="assistant"]').last();
    await expect(assistantMessage).toBeVisible();
    await assistantMessage.screenshot({
      path: 'test-results/working-box-completed.png',
    });
  });
});

