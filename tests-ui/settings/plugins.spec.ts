import { test, expect } from '../fixtures';

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Open the Plugin Hub by clicking the "Plugins" button in the header. */
async function openPluginHub(page: import('@playwright/test').Page) {
  await page.getByRole('button', { name: 'Plugins' }).first().click();
  await expect(page.getByText('Plugins', { exact: true }).first()).toBeVisible({ timeout: 5000 });
}

/** Navigate to the Upload tab inside the Plugin Hub. */
async function openUploadTab(page: import('@playwright/test').Page) {
  await openPluginHub(page);
  await page.getByRole('button', { name: 'Upload', exact: true }).click();
  await expect(page.getByText('Install local plugin', { exact: true }).first()).toBeVisible({ timeout: 3000 });
}

/** Navigate to the MCP Servers tab inside the Plugin Hub. */
async function openMcpServersTab(page: import('@playwright/test').Page) {
  await openPluginHub(page);
  await page.getByRole('button', { name: 'MCP Servers', exact: true }).click();
}

// ─── Plugin Hub navigation ───────────────────────────────────────────────────

test.describe('Plugin Hub', () => {
  test('opens from header Plugins button', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Plugins' }).first().click();
    await expect(page.getByText('Plugins', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });
  });

  test('closes via the Close button', async ({ configuredTaskpane: page }) => {
    await openPluginHub(page);
    await page.getByTitle('Close').click();
    await expect(page.getByRole('button', { name: 'Agent skills' })).toBeVisible();
  });

  test('shows Installed, Browse, MCP Servers, Upload and Marketplaces tabs', async ({
    configuredTaskpane: page,
  }) => {
    await openPluginHub(page);
    // Use exact: true to avoid matching "Installed (N)" collapsible section buttons
    await expect(page.getByRole('button', { name: 'Installed', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Browse', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'MCP Servers', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Upload', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Marketplaces', exact: true })).toBeVisible();
  });
});

// ─── Skill Picker (header toggle) ────────────────────────────────────────────

test.describe('Skill Picker', () => {
  test('Agent skills button is visible in header', async ({ configuredTaskpane: page }) => {
    await expect(page.getByRole('button', { name: 'Agent skills' })).toBeVisible({
      timeout: 5000,
    });
  });

  test('opens a popover listing bundled skills', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Agent skills' }).click();
    // The popover renders a "Skills" section label
    await expect(page.getByText('Skills', { exact: true }).first()).toBeVisible({ timeout: 3000 });
  });

  test('skills have aria-pressed attribute', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Agent skills' }).click();
    const skillToggles = page.locator('[aria-pressed]');
    await expect(skillToggles.first()).toBeVisible({ timeout: 3000 });
  });

  test('toggling a skill flips its aria-pressed state', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Agent skills' }).click();

    const enabledSkill = page.locator('[aria-pressed="true"]').first();
    await expect(enabledSkill).toBeVisible({ timeout: 3000 });

    await enabledSkill.click();

    await expect(page.locator('[aria-pressed="false"]').first()).toBeVisible({ timeout: 2000 });
  });

  test('Manage plugins link opens the Plugin Hub', async ({ configuredTaskpane: page }) => {
    await page.getByRole('button', { name: 'Agent skills' }).click();
    await page.getByRole('button', { name: 'Manage plugins…' }).click();
    await expect(page.getByText('Plugins', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });
  });
});

// ─── MCP Servers tab ─────────────────────────────────────────────────────────

test.describe('MCP Servers tab', () => {
  test('shows workiq and powerbi bundled servers', async ({ configuredTaskpane: page }) => {
    await openMcpServersTab(page);
    await expect(page.getByText('workiq')).toBeVisible({ timeout: 5000 });
    await expect(page.getByText('powerbi')).toBeVisible({ timeout: 5000 });
  });

  test('each server has a toggle button', async ({ configuredTaskpane: page }) => {
    await openMcpServersTab(page);
    await expect(page.getByRole('button', { name: 'Toggle workiq' })).toBeVisible({ timeout: 5000 });
    await expect(page.getByRole('button', { name: 'Toggle powerbi' })).toBeVisible({ timeout: 5000 });
  });

  test('toggle button flips aria-pressed state', async ({ configuredTaskpane: page }) => {
    await openMcpServersTab(page);

    const toggleBtn = page.getByRole('button', { name: 'Toggle workiq' });
    await expect(toggleBtn).toBeVisible({ timeout: 5000 });

    const initialPressed = await toggleBtn.getAttribute('aria-pressed');
    await toggleBtn.click();
    const newPressed = await toggleBtn.getAttribute('aria-pressed');
    expect(newPressed).not.toBe(initialPressed);
  });

  test('does not show a Remove or JSON upload button', async ({ configuredTaskpane: page }) => {
    await openMcpServersTab(page);
    await expect(page.getByText('workiq')).toBeVisible({ timeout: 5000 });
    await expect(page.getByTitle('Import MCP servers from JSON')).not.toBeVisible();
    await expect(page.getByRole('button', { name: /Remove/ })).not.toBeVisible();
  });
});

// ─── Upload tab ───────────────────────────────────────────────────────────────

test.describe('Upload tab', () => {
  test('shows "Install local plugin" heading', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);
    await expect(page.getByText('Install local plugin', { exact: true })).toBeVisible({ timeout: 3000 });
  });

  test('does not show Skills or Agents upload sections', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);
    // Old separate Skills/Agents upload panels are gone — only the plugin install form
    await expect(page.getByTitle('Import skills from ZIP')).not.toBeVisible();
    await expect(page.getByTitle('Import agents from ZIP')).not.toBeVisible();
    await expect(page.getByTitle('Import MCP servers from JSON')).not.toBeVisible();
  });
});
