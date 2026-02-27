/**
 * Playwright tests for the tool-fallback card UX.
 *
 * Uses the REAL Copilot API through the dev server (no mocks).
 * Requires `npm run dev` to be running on https://localhost:3000.
 *
 * Verifies:
 * - The trigger label shows only the humanized tool name (no "Used:", "Running:", "Cancelled:" prefixes)
 * - A result summary line appears in the trigger header after completion
 * - Expanding the card reveals "Input" and "Output" section headers
 * - report_intent does not produce a visible tool card
 */

import { test, expect } from '../fixtures';

const AI_TIMEOUT = 45_000;

/**
 * Wait for at least one completed tool card to appear in the thread and return
 * the first one. Uses manage_skills (action: list) which doesn't require Excel
 * so it runs reliably in any Playwright environment.
 */
async function waitForCompletedToolCard(page: import('@playwright/test').Page) {
  const composer = page.getByPlaceholder('Send a message...');
  await expect(composer).toBeVisible({ timeout: 10_000 });

  await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
  await expect(page.getByText('Connection failed')).not.toBeVisible();

  await composer.fill(
    'Use the manage_skills tool with action "list" and report back how many skills there are.'
  );
  await composer.press('Enter');

  // Wait for the response to finish
  await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({
    timeout: AI_TIMEOUT,
  });

  // At least one tool card must have rendered
  const card = page.locator('[data-slot="tool-fallback-root"]').first();
  await expect(card).toBeVisible({ timeout: 5_000 });
  return card;
}

test.describe('Tool card UX (live Copilot)', () => {
  test('trigger label shows only the tool name — no "Used:" / "Running:" / "Cancelled:" prefix', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    const trigger = card.locator('[data-slot="tool-fallback-trigger"]');

    // Must NOT contain legacy status prefixes
    await expect(trigger).not.toContainText(/\bUsed:\s/i);
    await expect(trigger).not.toContainText(/\bRunning:\s/i);
    await expect(trigger).not.toContainText(/\bCancelled:\s/i);

    // Must contain a human-readable tool name (the manage_skills label)
    await expect(trigger).toContainText(/manage skills/i);
  });

  test('trigger shows a result summary after tool completes', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    const trigger = card.locator('[data-slot="tool-fallback-trigger"]');

    // A non-empty summary span should appear between the name and the chevron
    const summary = trigger.locator('span.text-muted-foreground');
    await expect(summary).toBeVisible({ timeout: 5_000 });
    const text = await summary.textContent();
    expect(text?.trim().length).toBeGreaterThan(0);
  });

  test('expanded card shows "Input" header (not "Input:" or raw JSON unlabelled)', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);

    // Open the card
    await card.locator('[data-slot="tool-fallback-trigger"]').click();

    const args = card.locator('[data-slot="tool-fallback-args"]');
    await expect(args).toBeVisible({ timeout: 3_000 });

    // Must have "Input" as section header
    await expect(args).toContainText('Input');
    // Must NOT use old "Result:" naming
    await expect(args).not.toContainText('Result:');
  });

  test('expanded card shows "Output" header (not "Result:" old label)', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);

    // Open the card
    await card.locator('[data-slot="tool-fallback-trigger"]').click();

    const result = card.locator('[data-slot="tool-fallback-result"]');
    await expect(result).toBeVisible({ timeout: 3_000 });

    // Must use the new "Output" header
    await expect(result).toContainText('Output');
    // Must NOT use old "Result:" naming
    await expect(result).not.toContainText('Result:');
  });

  test('report_intent does not produce a tool card in the thread', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 10_000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });

    await composer.fill(
      'Use the manage_skills tool with action "list" and report back how many skills there are.'
    );
    await composer.press('Enter');

    await expect(page.getByRole('button', { name: 'Stop' })).not.toBeVisible({
      timeout: AI_TIMEOUT,
    });

    // No tool card should ever show "report_intent" as its label
    const toolCards = page.locator('[data-slot="tool-fallback-root"]');
    const count = await toolCards.count();
    for (let i = 0; i < count; i++) {
      await expect(toolCards.nth(i)).not.toContainText(/report.intent/i);
    }
  });
});
