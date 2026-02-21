import { test as base, type Page } from '@playwright/test';

/**
 * Polyfill for OfficeRuntime.storage using localStorage.
 * Must run before the app code so Zustand persist can hydrate.
 */
function officeRuntimePolyfill() {
  (globalThis as Record<string, unknown>).OfficeRuntime = {
    storage: {
      getItem: (key: string) => Promise.resolve(localStorage.getItem(key)),
      setItem: (key: string, value: string) => {
        localStorage.setItem(key, value);
        return Promise.resolve();
      },
      removeItem: (key: string) => {
        localStorage.removeItem(key);
        return Promise.resolve();
      },
    },
  };
}

/** Minimal settings blob matching the current UserSettings shape. */
function makeSettingsJSON(overrides: Record<string, unknown> = {}) {
  return JSON.stringify({
    state: {
      activeModel: 'claude-sonnet-4',
      activeSkillNames: null,
      activeAgentId: 'Excel',
      importedSkills: [],
      importedAgents: [],
      importedMcpServers: [],
      activeMcpServerNames: null,
      availableModels: [
        { id: 'claude-sonnet-4', name: 'Claude Sonnet 4', provider: 'Anthropic' },
        { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
        { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro', provider: 'Google' },
      ],
      ...overrides,
    },
  });
}

// ─── Mock WebSocket server helpers ───────────────────────────────────────────
// The app uses LSP framing over WebSocket: `Content-Length: N\r\n\r\n<json>`
// These helpers parse incoming messages and send correctly framed responses.

/**
 * Models returned by the mock WebSocket server.
 * Deliberately different from the app's default model (claude-sonnet-4) to
 * exercise the auto-correction path in loadAvailableModels().
 */
export const MOCK_SERVER_MODELS = [
  { id: 'mock-model-opus', name: 'Mock Model Opus' },
  { id: 'mock-model-turbo', name: 'Mock Model Turbo' },
];

/** Parse an LSP-framed JSON-RPC message from the WebSocket transport. */
function parseLspMessage(
  raw: string | Buffer
): { id: number; method?: string; params?: unknown } | null {
  const text = typeof raw === 'string' ? raw : raw.toString('utf-8');
  const idx = text.indexOf('\r\n\r\n');
  if (idx === -1) return null;
  try {
    return JSON.parse(text.slice(idx + 4)) as { id: number; method?: string; params?: unknown };
  } catch {
    return null;
  }
}

/** Wrap a JSON-RPC result in an LSP-framed response string. */
function makeLspResponse(id: number, result: unknown): string {
  const json = JSON.stringify({ jsonrpc: '2.0', id, result });
  const byteLen = Buffer.byteLength(json, 'utf-8');
  return `Content-Length: ${byteLen}\r\n\r\n${json}`;
}

/** Wrap a JSON-RPC error in an LSP-framed response string. */
function makeLspError(id: number, code: number, message: string): string {
  const json = JSON.stringify({ jsonrpc: '2.0', id, error: { code, message } });
  const byteLen = Buffer.byteLength(json, 'utf-8');
  return `Content-Length: ${byteLen}\r\n\r\n${json}`;
}

// ─── Fixtures ─────────────────────────────────────────────────────────────────
/**
 * Shared fixtures for UI tests.
 *
 * - `taskpane`: fresh state, no pre-seeded data, no WS mock (real dev server).
 * - `configuredTaskpane`: pre-seeds localStorage with known model/agent/skill
 *    settings. Use for deterministic UI rendering tests that don't test the
 *    connection flow.
 * - `mockServerTaskpane`: fresh state with a mock WebSocket server that
 *    responds to session.create and models.list. Tests the real connection
 *    flow — models are genuinely fetched and applied to the UI.
 * - `disconnectedTaskpane`: WS connection is accepted then immediately closed.
 *    Tests that the app correctly surfaces session errors to the user.
 */
export const test = base.extend<{
  taskpane: Page;
  configuredTaskpane: Page;
  mockServerTaskpane: Page;
  disconnectedTaskpane: Page;
}>({
  /** Navigate to the task pane (default/fresh state). */
  taskpane: async ({ page }, use) => {
    await page.addInitScript(officeRuntimePolyfill);
    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },

  /**
   * Navigate with pre-seeded settings for deterministic UI rendering tests.
   * availableModels is seeded directly — no WS connection is made.
   * Use this for testing UI components (model picker, agent picker, etc.)
   * with a known, stable data set.
   */
  configuredTaskpane: async ({ page }, use) => {
    await page.addInitScript(officeRuntimePolyfill);
    await page.addInitScript((json: string) => {
      localStorage.setItem('office-coding-agent-settings', json);
    }, makeSettingsJSON());
    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },

  /**
   * Navigate with a mock WebSocket server that speaks the app's JSON-RPC
   * protocol. Tests the REAL connection flow: the app connects, sends
   * session.create and models.list, and the UI updates from the responses.
   *
   * MOCK_SERVER_MODELS contains IDs that do NOT match the default activeModel
   * (claude-sonnet-4), deliberately triggering the auto-correction path.
   */
  mockServerTaskpane: async ({ page }, use) => {
    await page.addInitScript(officeRuntimePolyfill);

    // Intercept the WebSocket before navigating so the mock is in place
    await page.routeWebSocket('wss://localhost:3000/api/copilot', ws => {
      ws.onMessage(raw => {
        const msg = parseLspMessage(raw);
        if (!msg || msg.id === undefined) return;

        if (msg.method === 'session.create') {
          ws.send(makeLspResponse(msg.id, { sessionId: 'mock-session-1' }));
        } else if (msg.method === 'models.list') {
          ws.send(makeLspResponse(msg.id, { models: MOCK_SERVER_MODELS }));
        }
        // Other methods (session.destroy, etc.) are silently ignored
      });
    });

    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },

  /**
   * Navigate with a WebSocket mock that responds to every JSON-RPC request
   * with a server-error response, simulating a server that is reachable but
   * rejects all operations (e.g., authentication failure, service down).
   *
   * The sequence: WebSocket opens → session.create receives an error response
   * → createSession() rejects with ResponseError → sessionError is set →
   * SessionErrorBanner renders. availableModels is never populated (because
   * loadAvailableModels() is only called after a successful createSession()),
   * so the model picker shows "Connecting to Copilot…".
   *
   * This approach is deterministic because it uses the normal JSON-RPC error
   * response path rather than relying on WebSocket close timing.
   */
  disconnectedTaskpane: async ({ page }, use) => {
    await page.addInitScript(officeRuntimePolyfill);

    await page.routeWebSocket('wss://localhost:3000/api/copilot', ws => {
      ws.onMessage(raw => {
        const msg = parseLspMessage(raw);
        if (!msg || msg.id === undefined) return;
        // Reject every request with a server-side error
        ws.send(makeLspError(msg.id, -32001, 'Server is unavailable'));
      });
    });

    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },
});

export { expect } from '@playwright/test';
