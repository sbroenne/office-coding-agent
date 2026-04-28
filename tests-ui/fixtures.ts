import { test as base, type APIRequestContext, type Page } from '@playwright/test';

/**
 * Polyfill for OfficeRuntime.storage using localStorage.
 * Must run before the app code so Zustand persist can hydrate.
 */
function officeRuntimePolyfill() {
  const officeHost = 'excel';
  (globalThis as Record<string, unknown>).Office = {
    HostType: {
      Excel: 'excel',
      PowerPoint: 'powerpoint',
      Word: 'word',
    },
    context: {
      host: officeHost,
    },
    onReady: (callback?: () => void) => {
      if (typeof callback === 'function') callback();
      return Promise.resolve({ host: officeHost });
    },
  };

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
      disabledSkillNames: [],
      disabledMcpServerNames: [],
      availableModels: [
        { id: 'claude-sonnet-4', name: 'Claude Sonnet 4', provider: 'Anthropic' },
        { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
        { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro', provider: 'Google' },
      ],
      ...overrides,
    },
  });
}

export async function hasDevServer(request: APIRequestContext): Promise<boolean> {
  try {
    const response = await request.get('/api/ping');
    return response.ok();
  } catch {
    return false;
  }
}

/**
 * Shared fixtures for UI tests.
 *
 * These fixtures never intercept WebSocket traffic. Playwright coverage in this
 * repo must exercise the real proxy server and Copilot session flow.
 */
export const test = base.extend<{
  taskpane: Page;
  configuredTaskpane: Page;
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
   * This only seeds persisted settings; the page still talks to the live proxy.
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
});

export { expect } from '@playwright/test';
