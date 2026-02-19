import { test as base, type Page } from '@playwright/test';

/**
 * Shared fixtures for UI tests.
 *
 * - `taskpane`: navigates to the task pane (fresh state).
 * - `configuredTaskpane`: pre-seeds localStorage with known settings for
 *    deterministic UI state (active model, agent, skills).
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
  /** Navigate to the task pane (fresh/default state). */
  taskpane: async ({ page }, use) => {
    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },

  /** Navigate with pre-seeded settings for deterministic UI state. */
  configuredTaskpane: async ({ page }, use) => {
    // Seed the Zustand persisted store BEFORE navigating
    await page.addInitScript((json: string) => {
      localStorage.setItem('office-coding-agent-settings', json);
    }, makeSettingsJSON());

    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },
});

export { expect } from '@playwright/test';
