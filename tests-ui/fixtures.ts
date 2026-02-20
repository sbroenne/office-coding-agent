import { test as base, type Page } from '@playwright/test';

/**
 * Shared fixtures for UI tests.
 *
 * - `taskpane`: navigates to the task pane (fresh state).
 * - `configuredTaskpane`: pre-seeds localStorage with known settings for
 *    deterministic UI state (active model, agent, skills).
 */

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

/**
 * Shared fixtures for UI tests.
 *
 * - `taskpane`: navigates to the task pane with default settings.
 * - `configuredTaskpane`: pre-seeds localStorage with known settings for
 *    deterministic UI state (active model, agent, skills).
 *
 * Both fixtures inject an OfficeRuntime.storage polyfill so Zustand persist
 * can hydrate from localStorage in the Chromium test browser.
 */

/** Minimal settings blob matching the current UserSettings shape. */
function makeSettingsJSON(overrides: Record<string, unknown> = {}) {
  return JSON.stringify({
    state: {
      activeModel: 'claude-sonnet-4.5',
      activeSkillNames: null,
      activeAgentId: 'Excel',
      importedSkills: [],
      importedAgents: [],
      importedMcpServers: [],
      activeMcpServerNames: null,
      ...overrides,
    },
    version: 0,
  });
}

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

  /** Navigate with pre-seeded settings for deterministic UI state. */
  configuredTaskpane: async ({ page }, use) => {
    // Inject OfficeRuntime polyfill FIRST so Zustand persist can hydrate
    await page.addInitScript(officeRuntimePolyfill);
    // Then seed the Zustand persisted store
    await page.addInitScript((json: string) => {
      localStorage.setItem('office-coding-agent-settings', json);
    }, makeSettingsJSON());

    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },
});

export { expect } from '@playwright/test';
