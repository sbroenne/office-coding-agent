/**
 * Playwright tests for the tool-fallback card UX.
 *
 * Prerequisite: `npm run dev` must be serving https://localhost:3000.
 * These scenarios exercise the real proxy + Copilot session by asking the model
 * to call bundled management tools. No mocked WebSocket traffic is allowed.
 */

import { test, expect, hasDevServer } from '../fixtures';

const AI_TIMEOUT = 45_000;

async function waitForCompletedToolCard(page: import('@playwright/test').Page) {
  const composer = page.getByPlaceholder('Send a message...');
  await expect(composer).toBeVisible({ timeout: 10_000 });
  await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({ timeout: 15_000 });
  await expect(page.getByText('Connection failed')).not.toBeVisible();

  await composer.fill(
    'Use the manage_plugins tool with action "list". After the tool completes, briefly summarize the result.'
  );
  await composer.press('Enter');

  const toolCard = page.locator('[data-slot="tool-fallback-root"]').first();
  const workingBox = page.locator('.chat-thinking-box').first();
  await expect
    .poll(
      async () =>
        (await toolCard.isVisible().catch(() => false)) ||
        (await workingBox.isVisible().catch(() => false)),
      { timeout: AI_TIMEOUT }
    )
    .toBe(true);

  await expect(page.getByRole('button', { name: /^(Stop|Cancel)$/ })).not.toBeVisible({
    timeout: AI_TIMEOUT,
  });

  if (
    !(await toolCard.isVisible().catch(() => false)) &&
    (await workingBox.isVisible().catch(() => false))
  ) {
    await workingBox
      .locator('.chat-thinking-header')
      .click()
      .catch(() => undefined);
  }

  await expect(toolCard).toBeVisible({ timeout: 10_000 });
  return toolCard;
}

test.describe('Tool card UX (live Copilot)', () => {
  test.beforeEach(async ({ request }) => {
    test.skip(
      !(await hasDevServer(request)),
      'tests-ui requires `npm run dev` with the live proxy server on https://localhost:3000'
    );
  });

  test('trigger label shows only the tool name — no "Used:" / "Running:" / "Cancelled:" prefix', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    const trigger = card.locator('[data-slot="tool-fallback-trigger"]');

    await expect(trigger).not.toContainText(/\bUsed:\s/i);
    await expect(trigger).not.toContainText(/\bRunning:\s/i);
    await expect(trigger).not.toContainText(/\bCancelled:\s/i);
    await expect(trigger).toContainText(/manage plugins/i);
  });

  test('trigger shows a result summary after tool completes', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    const trigger = card.locator('[data-slot="tool-fallback-trigger"]');
    const summary = trigger.locator('.progress-summary');

    await expect(summary).toBeVisible({ timeout: 10_000 });
    const text = await summary.textContent();
    expect(text?.trim().length).toBeGreaterThan(0);
  });

  test('expanded card shows "Input" header', async ({ configuredTaskpane: page }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    await card.locator('[data-slot="tool-fallback-trigger"]').click();

    const details = card.locator('.tool-details-expanded');
    await expect(details).toBeVisible({ timeout: 10_000 });
    await expect(details).toContainText('Input');
  });

  test('expanded card shows "Output" header and not the legacy "Result:" label', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    await card.locator('[data-slot="tool-fallback-trigger"]').click();

    const details = card.locator('.tool-details-expanded');
    await expect(details).toBeVisible({ timeout: 10_000 });
    await expect(details).toContainText('Output');
    await expect(details).not.toContainText('Result:');
  });

  test('report_intent does not produce a tool card in the thread', async ({
    configuredTaskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const card = await waitForCompletedToolCard(page);
    await expect(card).not.toContainText(/report.intent/i);

    const toolCards = page.locator('[data-slot="tool-fallback-root"]');
    const count = await toolCards.count();
    for (let i = 0; i < count; i++) {
      await expect(toolCards.nth(i)).not.toContainText(/report.intent/i);
    }
  });
});
