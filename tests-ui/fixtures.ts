import { test as base, type Page } from '@playwright/test';

/**
 * Shared fixtures for UI tests.
 *
 * - `taskpane`: navigates to the task pane and waits for the app to render.
 * - `configuredTaskpane`: pre-seeds localStorage with a valid settings store
 *    so the app skips the SetupWizard and renders the chat UI immediately.
 */

/** Minimal settings blob that passes the `needsSetup` check in App.tsx */
function makeSettingsJSON(overrides: Record<string, unknown> = {}) {
  return JSON.stringify({
    state: {
      endpoints: [
        {
          id: 'ep-test',
          displayName: 'Test Endpoint',
          resourceUrl: 'https://test.openai.azure.com',
          authMethod: 'apiKey',
          apiKey: 'test-key-for-ui-tests',
        },
      ],
      activeEndpointId: 'ep-test',
      activeModelId: 'gpt-4.1',
      defaultModelId: 'gpt-4.1',
      endpointModels: {
        'ep-test': [
          { id: 'gpt-4.1', name: 'gpt-4.1', ownedBy: 'azure', provider: 'OpenAI' },
          { id: 'gpt-5.2-chat', name: 'gpt-5.2-chat', ownedBy: 'azure', provider: 'OpenAI' },
        ],
      },
      activeSkillNames: null,
      activeAgentId: 'Excel',
      importedSkills: [],
      importedAgents: [],
      ...overrides,
    },
    version: 0,
  });
}

export const test = base.extend<{
  taskpane: Page;
  configuredTaskpane: Page;
}>({
  /** Navigate to the task pane (fresh state — will show SetupWizard) */
  taskpane: async ({ page }, use) => {
    await page.goto('/taskpane.html');
    await page.waitForLoadState('domcontentloaded');
    await use(page);
  },

  /** Navigate with pre-seeded settings (skips wizard → shows chat UI) */
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
