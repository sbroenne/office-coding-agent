/**
 * Tests for the app's real WebSocket connection flow using a mock WS server.
 *
 * Unlike `configuredTaskpane` tests (which bypass the connection by pre-seeding
 * localStorage), these tests use the `mockServerTaskpane` fixture which:
 *   1. Intercepts the WebSocket via page.routeWebSocket()
 *   2. Responds to session.create and models.list with deterministic data
 *   3. Starts with NO pre-seeded availableModels
 *
 * This validates the complete pipeline:
 *   WS connect → session.create → models.list → store update → UI render
 *
 * The mock server returns MOCK_SERVER_MODELS (IDs that do NOT match the default
 * activeModel), deliberately exercising the auto-correction path in
 * loadAvailableModels().
 */

import { test, expect, MOCK_SERVER_MODELS } from '../fixtures';

test.describe('Chat UI — mock WebSocket server (real connection flow)', () => {
  test('fetches models from the server and auto-corrects the active model', async ({
    mockServerTaskpane: page,
  }) => {
    // The default activeModel ('claude-sonnet-4') is not in MOCK_SERVER_MODELS.
    // loadAvailableModels() auto-corrects to models[0] = 'mock-model-opus'.
    // The model picker button must therefore show 'Mock Model Opus'.
    await expect(page.getByText(MOCK_SERVER_MODELS[0].name)).toBeVisible({ timeout: 10_000 });
  });

  test('model picker dropdown lists all models returned by the server', async ({
    mockServerTaskpane: page,
  }) => {
    // Wait for the auto-corrected model to appear in the toolbar button
    await expect(page.getByText(MOCK_SERVER_MODELS[0].name)).toBeVisible({ timeout: 10_000 });

    // Open the model picker (toolbar button — aria-label="Select model")
    await page.getByRole('button', { name: 'Select model' }).click();

    // Both models from the mock server must appear as option buttons in the
    // dropdown. Use getByRole('button', { name }) to target the dropdown
    // buttons (accessible name = model name) rather than getByText(), which
    // would match the toolbar button span too (strict mode violation).
    for (const model of MOCK_SERVER_MODELS) {
      await expect(page.getByRole('button', { name: model.name })).toBeVisible({ timeout: 3_000 });
    }
  });
});
