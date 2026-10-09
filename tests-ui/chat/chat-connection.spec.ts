/**
 * Playwright tests for the real connection flow.
 *
 * Prerequisite: `npm run dev` must be serving https://localhost:3000.
 * These checks intentionally use the live proxy session and never mock WebSocket
 * traffic or inject synthetic JSON-RPC events.
 */

import { test, expect, hasDevServer } from '../fixtures';

const CONNECTION_TIMEOUT = 15_000;
const AI_TIMEOUT = 45_000;

test.describe('Chat UI — live connection flow', () => {
  test('starts interactive and switches Plan and Autopilot through the live session', async ({
    taskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({
      timeout: CONNECTION_TIMEOUT,
    });
    const picker = page.getByRole('button', { name: 'Select conversation mode', exact: true });
    await expect(picker).toHaveText('Interactive');
    for (const mode of ['Plan', 'Autopilot', 'Interactive']) {
      await picker.click();
      const selectedMode = (await picker.textContent()) ?? '';
      await expect(page.getByRole('button', { name: selectedMode, exact: true })).toHaveAttribute(
        'aria-pressed',
        'true'
      );
      await page.getByRole('button', { name: mode, exact: true }).click();
      await expect(picker).toHaveText(mode);
    }
    await picker.click();
    await page.getByRole('button', { name: 'Autopilot', exact: true }).click();
    await expect(picker).toHaveText('Autopilot');
    await page
      .getByRole('textbox', { name: 'Message input' })
      .fill('Reply with exactly one word: READY');
    await page.getByRole('button', { name: 'Send', exact: true }).click();
    await expect(page.locator('[data-role="assistant"]').last()).toContainText(/ready/i, {
      timeout: AI_TIMEOUT,
    });
    await page.getByRole('button', { name: 'New conversation', exact: true }).click();
    await expect(picker).toHaveText('Interactive');
  });

  test.beforeEach(async ({ request }) => {
    test.skip(
      !(await hasDevServer(request)),
      'tests-ui requires `npm run dev` with the live proxy server on https://localhost:3000'
    );
  });

  test('fresh taskpane loads live model options', async ({ taskpane: page }) => {
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({
      timeout: CONNECTION_TIMEOUT,
    });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    const modelButton = page.getByRole('button', { name: 'Select model' });
    await expect(modelButton).toBeVisible({ timeout: 10_000 });
    await modelButton.click();

    const modelOptions = page
      .getByRole('button')
      .filter({ hasText: /(Claude|GPT|Gemini|o[0-9])/i });
    await expect(modelOptions.first()).toBeVisible({ timeout: 10_000 });
  });

  test('fresh taskpane can send a message after the live session connects', async ({
    taskpane: page,
  }) => {
    test.setTimeout(AI_TIMEOUT + 30_000);

    const composer = page.getByPlaceholder('Send a message...');
    await expect(composer).toBeVisible({ timeout: 10_000 });
    await expect(page.getByText('Connecting to Copilot...')).not.toBeVisible({
      timeout: CONNECTION_TIMEOUT,
    });
    await expect(page.getByText('Connection failed')).not.toBeVisible();

    await composer.fill('Reply with exactly one word: READY');
    await composer.press('Enter');

    await expect(page.getByText(/ready/i).first()).toBeVisible({ timeout: AI_TIMEOUT });
  });
});
