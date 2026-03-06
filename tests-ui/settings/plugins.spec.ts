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
  await expect(page.getByText('Skills', { exact: true }).first()).toBeVisible({ timeout: 3000 });
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

  test('shows Installed, Browse, Marketplaces, and Upload tabs', async ({
    configuredTaskpane: page,
  }) => {
    await openPluginHub(page);
    // Use exact: true to avoid matching "Installed (N)" collapsible section buttons
    await expect(page.getByRole('button', { name: 'Installed', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Browse', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Marketplaces', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Upload', exact: true })).toBeVisible();
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

// ─── Upload tab — Skills ──────────────────────────────────────────────────────

test.describe('Upload tab — Skills', () => {
  test('shows upload buttons for Skills section', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);
    await expect(page.getByTitle('Import skills from ZIP')).toBeVisible();
    await expect(page.getByTitle('Import a single skill .md file')).toBeVisible();
  });

  test('shows Uploaded count label', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);
    await expect(page.getByText('Uploaded (0)').first()).toBeVisible({ timeout: 3000 });
  });

  test('uploading a skill .md file adds it to the Uploaded list', async ({
    configuredTaskpane: page,
  }) => {
    await openUploadTab(page);

    const mdContent = [
      '---',
      'name: test-skill',
      'description: A test skill',
      'version: 1.0.0',
      'hosts: []',
      '---',
      '',
      'Do something useful.',
    ].join('\n');

    await page
      .locator('input[aria-label="Import skill Markdown file"]')
      .setInputFiles({ name: 'test-skill.md', mimeType: 'text/markdown', buffer: Buffer.from(mdContent) });

    await expect(page.getByText('test-skill', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });
  });

  test('Remove button removes an uploaded skill', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);

    const mdContent = [
      '---',
      'name: removable-skill',
      'description: Remove me',
      'version: 1.0.0',
      'hosts: []',
      '---',
      '',
      'To be removed.',
    ].join('\n');

    await page
      .locator('input[aria-label="Import skill Markdown file"]')
      .setInputFiles({ name: 'removable-skill.md', mimeType: 'text/markdown', buffer: Buffer.from(mdContent) });

    await expect(page.getByText('removable-skill', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });

    await page.getByRole('button', { name: 'Remove removable-skill' }).click();
    await expect(page.getByRole('button', { name: 'Remove removable-skill' })).not.toBeVisible({
      timeout: 3000,
    });
  });
});

// ─── Upload tab — Agents ─────────────────────────────────────────────────────

test.describe('Upload tab — Agents', () => {
  test('shows upload buttons for Agents section', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);
    await expect(page.getByText('Agents', { exact: true }).first()).toBeVisible({ timeout: 3000 });
    await expect(page.getByTitle('Import agents from ZIP')).toBeVisible();
    await expect(page.getByTitle('Import a single agent .md file')).toBeVisible();
  });

  test('uploading an agent .md file adds it to the Uploaded list', async ({
    configuredTaskpane: page,
  }) => {
    await openUploadTab(page);

    const mdContent = [
      '---',
      'name: test-agent',
      'description: A test agent',
      'version: 1.0.0',
      'hosts: [excel]',
      'defaultForHosts: []',
      '---',
      '',
      'Do something useful.',
    ].join('\n');

    await page
      .locator('input[aria-label="Import agent Markdown file"]')
      .setInputFiles({ name: 'test-agent.md', mimeType: 'text/markdown', buffer: Buffer.from(mdContent) });

    await expect(page.getByText('test-agent', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });
  });
});

// ─── Upload tab — MCP Servers ─────────────────────────────────────────────────

test.describe('Upload tab — MCP Servers', () => {
  test('shows upload button for MCP Servers section', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);
    await expect(page.getByText('MCP Servers', { exact: true })).toBeVisible({ timeout: 3000 });
    await expect(page.getByTitle('Import MCP servers from JSON')).toBeVisible();
  });

  test('uploading a JSON file adds a server to the list', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);

    const serverJson = JSON.stringify({ name: 'test-mcp', transport: 'http', url: 'https://example.com/mcp' });

    await page
      .locator('input[aria-label="Import MCP servers from JSON file"]')
      .setInputFiles({ name: 'test-mcp.json', mimeType: 'application/json', buffer: Buffer.from(serverJson) });

    await expect(page.getByText('test-mcp', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });
  });

  test('MCP toggle button flips aria-pressed state', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);

    const serverJson = JSON.stringify({ name: 'toggle-mcp', transport: 'http', url: 'https://example.com/mcp' });

    await page
      .locator('input[aria-label="Import MCP servers from JSON file"]')
      .setInputFiles({ name: 'toggle-mcp.json', mimeType: 'application/json', buffer: Buffer.from(serverJson) });

    await expect(page.getByText('toggle-mcp', { exact: true }).first()).toBeVisible({ timeout: 5000 });

    const toggleBtn = page.getByRole('button', { name: 'Toggle toggle-mcp' });
    const initialPressed = await toggleBtn.getAttribute('aria-pressed');
    await toggleBtn.click();
    expect(await toggleBtn.getAttribute('aria-pressed')).not.toBe(initialPressed);
  });

  test('Remove button removes an uploaded MCP server', async ({ configuredTaskpane: page }) => {
    await openUploadTab(page);

    const serverJson = JSON.stringify({ name: 'removable-mcp', transport: 'http', url: 'https://example.com/mcp' });

    await page
      .locator('input[aria-label="Import MCP servers from JSON file"]')
      .setInputFiles({ name: 'removable-mcp.json', mimeType: 'application/json', buffer: Buffer.from(serverJson) });

    await expect(page.getByText('removable-mcp', { exact: true }).first()).toBeVisible({
      timeout: 5000,
    });

    await page.getByRole('button', { name: 'Remove removable-mcp' }).click();
    await expect(page.getByRole('button', { name: 'Remove removable-mcp' })).not.toBeVisible({
      timeout: 3000,
    });
  });
});
