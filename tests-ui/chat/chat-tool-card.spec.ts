/**
 * Playwright tests for the tool-fallback card UX.
 *
 * Uses a deterministic mock WebSocket server (toolCardMockTaskpane fixture)
 * that injects synthetic tool call events on every message send. This makes
 * the tests fast (~300 ms per scenario) and removes any dependence on live
 * LLM calls or LLM non-determinism.
 *
 * Verifies:
 * - The trigger label shows only the humanized tool name (no "Used:", "Running:", "Cancelled:" prefixes)
 * - A result summary line appears in the trigger header after completion
 * - Expanding the card reveals "Input" and "Output" section headers
 * - report_intent does not produce a visible tool card
 */

import { test, expect } from '../fixtures';

/**
 * Send a message and wait for the mock server's synthetic manage_skills tool
 * call to produce a completed tool card. Returns the first tool card found.
 *
 * The mock fires: report_intent → manage_skills start → manage_skills complete
 * → text delta → session.idle, all within ~250 ms.
 */
async function waitForCompletedToolCard(page: import('@playwright/test').Page) {
  const composer = page.getByPlaceholder('Send a message...');
  await expect(composer).toBeVisible({ timeout: 10_000 });

  await composer.fill('List available skills.');
  await composer.press('Enter');

  // Working box appears as soon as manage_skills tool.execution_start fires (~80 ms)
  const workingBox = page.locator('.chat-thinking-box').first();
  await expect(workingBox).toBeVisible({ timeout: 5_000 });

  // Wait for the response to finish (Stop button disappears after session.idle)
  await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({ timeout: 5_000 });

  // The Working box auto-collapses after completion — expand it to access tool cards
  await workingBox.locator('.chat-thinking-header').click();

  // Tool card is now visible inside the expanded Working box
  const card = page.locator('[data-slot="tool-fallback-root"]').first();
  await expect(card).toBeVisible({ timeout: 3_000 });
  return card;
}

test.describe('Tool card UX (mock server)', () => {
  test('trigger label shows only the tool name — no "Used:" / "Running:" / "Cancelled:" prefix', async ({
    toolCardMockTaskpane: page,
  }) => {
    const card = await waitForCompletedToolCard(page);
    const trigger = card.locator('[data-slot="tool-fallback-trigger"]');

    // Must NOT contain legacy status prefixes
    await expect(trigger).not.toContainText(/\bUsed:\s/i);
    await expect(trigger).not.toContainText(/\bRunning:\s/i);
    await expect(trigger).not.toContainText(/\bCancelled:\s/i);

    // Must contain a human-readable tool name (manage_skills → "Manage skills")
    await expect(trigger).toContainText(/manage skills/i);
  });

  test('trigger shows a result summary after tool completes', async ({
    toolCardMockTaskpane: page,
  }) => {
    const card = await waitForCompletedToolCard(page);
    const trigger = card.locator('[data-slot="tool-fallback-trigger"]');

    // A non-empty summary span should appear (VS Code progress-summary class)
    const summary = trigger.locator('.progress-summary');
    await expect(summary).toBeVisible({ timeout: 5_000 });
    const text = await summary.textContent();
    expect(text?.trim().length).toBeGreaterThan(0);
  });

  test('expanded card shows "Input" header (not "Input:" or raw JSON unlabelled)', async ({
    toolCardMockTaskpane: page,
  }) => {
    const card = await waitForCompletedToolCard(page);

    // Open the card by clicking the trigger
    await card.locator('[data-slot="tool-fallback-trigger"]').click();

    const details = card.locator('.tool-details-expanded');
    await expect(details).toBeVisible({ timeout: 3_000 });

    // Must have "Input" as section header
    await expect(details).toContainText('Input');
  });

  test('expanded card shows "Output" header (not "Result:" old label)', async ({
    toolCardMockTaskpane: page,
  }) => {
    const card = await waitForCompletedToolCard(page);

    // Open the card
    await card.locator('[data-slot="tool-fallback-trigger"]').click();

    const details = card.locator('.tool-details-expanded');
    await expect(details).toBeVisible({ timeout: 3_000 });

    // Must use the new "Output" header
    await expect(details).toContainText('Output');
    // Must NOT use old "Result:" naming
    await expect(details).not.toContainText('Result:');
  });

  test('report_intent does not produce a tool card in the thread', async ({
    toolCardMockTaskpane: page,
  }) => {
    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 10_000 });

    await composer.fill('List available skills.');
    await composer.press('Enter');

    // Wait for the response to finish
    await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({ timeout: 5_000 });

    // Expand the Working box (mock fires report_intent first)
    const workingBox = page.locator('.chat-thinking-box').first();
    await expect(workingBox).toBeVisible({ timeout: 3_000 });
    await workingBox.locator('.chat-thinking-header').click();

    // No tool card should show "report_intent" as its label
    const toolCards = page.locator('[data-slot="tool-fallback-root"]');
    const count = await toolCards.count();
    for (let i = 0; i < count; i++) {
      await expect(toolCards.nth(i)).not.toContainText(/report.intent/i);
    }
  });
});

