/**
 * Tests for the app's disconnected/error state.
 *
 * Uses the `disconnectedTaskpane` fixture which intercepts the WebSocket and
 * immediately closes it without responding to any requests. This causes
 * session.create() to fail, setting sessionError and triggering
 * SessionErrorBanner plus the model picker "Connecting to Copilot…" state.
 *
 * These are the tests that would have caught the bugs fixed in the session:
 *  - sessionError returned by useOfficeChat but never rendered in App.tsx
 *  - Model picker showing an empty dropdown instead of "Connecting to Copilot…"
 *  - onNew() silently returning when sessionRef.current is null
 */

import { test, expect } from '../fixtures';

// All disconnected tests share the same WS-failure setup — allow more time
// for the session error to propagate through the async WS/JSON-RPC stack.
const SESSION_ERROR_TIMEOUT = 15_000;

test.describe('Chat UI — disconnected state (server unavailable)', () => {
  test('shows session error banner when the server closes the connection', async ({
    disconnectedTaskpane: page,
  }) => {
    await expect(page.getByText(/Connection failed:/)).toBeVisible({
      timeout: SESSION_ERROR_TIMEOUT,
    });
    await expect(page.getByRole('button', { name: 'Retry' })).toBeVisible();
  });

  test('model picker dropdown shows "Connecting to Copilot…" when no models are available', async ({
    disconnectedTaskpane: page,
  }) => {
    // Wait until the session error is confirmed (no models were fetched)
    await expect(page.getByText(/Connection failed:/)).toBeVisible({
      timeout: SESSION_ERROR_TIMEOUT,
    });

    // Open the model picker — it should show the loading/connecting state
    await page.getByRole('button', { name: 'Select model' }).click();
    await expect(page.getByText(/Connecting to Copilot/)).toBeVisible({ timeout: 3_000 });
  });

  test('sending a message with no session shows a "Not connected" error in the chat', async ({
    disconnectedTaskpane: page,
  }) => {
    // Confirm the session failed before attempting to send
    await expect(page.getByText(/Connection failed:/)).toBeVisible({
      timeout: SESSION_ERROR_TIMEOUT,
    });

    // Type and submit a message
    await page.getByPlaceholder('Send a message...').fill('Hello');
    await page.keyboard.press('Enter');

    // The assistant should respond with a "Not connected" error message
    await expect(page.getByText(/Not connected to Copilot/)).toBeVisible({ timeout: 5_000 });
  });
});
