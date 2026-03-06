/**
 * Regression tests: PluginHub — Marketplace tab (delete & edit).
 *
 * These tests document the expected behaviour for the Marketplaces tab:
 * - Delete button is visible for custom marketplaces that have a registeredKey
 * - Delete button is hidden for the OCA marketplace (isOwn = true)
 * - Delete button is hidden for built-in marketplaces (isBuiltIn = true)
 * - Delete button is hidden when registeredKey is null (CLI mapping failure)
 * - Clicking delete calls removeMarketplace with the registeredKey (not slug)
 * - Lock icon is shown for protected marketplaces
 * - After a successful removal the list is refreshed
 */
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { renderWithProviders } from '../test-utils';
import PluginHub from '@/components/plugins/PluginHub';
import type { RegisteredMarketplace } from '@/types/plugin';
import * as pluginService from '@/services/plugins/pluginService';

// ─── Helpers ──────────────────────────────────────────────────────────────────

function makeMarketplace(overrides: Partial<RegisteredMarketplace>): RegisteredMarketplace {
  return {
    slug: 'sbroenne-my-plugins',
    name: 'My Plugins',
    source: 'sbroenne/my-plugins',
    isBuiltIn: false,
    isOwn: false,
    registeredKey: 'sbroenne/my-plugins',
    pluginCount: 3,
    ...overrides,
  };
}

/** Navigate the rendered PluginHub to the Marketplaces tab. */
async function openMarketplacesTab(user: ReturnType<typeof userEvent.setup>) {
  const tab = await screen.findByRole('button', { name: /marketplaces/i });
  await user.click(tab);
}

// ─── Setup ────────────────────────────────────────────────────────────────────

beforeEach(() => {
  // Default: each service call returns empty lists / success
  vi.spyOn(pluginService, 'getInstalledPlugins').mockResolvedValue([]);
  vi.spyOn(pluginService, 'getMarketplaces').mockResolvedValue([]);
  vi.spyOn(pluginService, 'browseMarketplace').mockResolvedValue([]);
  vi.spyOn(pluginService, 'removeMarketplace').mockResolvedValue({ success: true, message: 'removed' });
  vi.spyOn(pluginService, 'addMarketplace').mockResolvedValue({ success: true, message: 'added' });
});

afterEach(() => {
  vi.restoreAllMocks();
});

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('Integration: PluginHub — Marketplaces tab', () => {
  it('delete button is visible for a custom marketplace with a registeredKey', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([
      makeMarketplace({ registeredKey: 'sbroenne/my-plugins', isOwn: false, isBuiltIn: false }),
    ]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    const deleteBtn = await screen.findByRole('button', { name: /remove marketplace/i });
    expect(deleteBtn).toBeInTheDocument();
  });

  it('delete button is HIDDEN when registeredKey is null (CLI slug mismatch bug)', async () => {
    // Regression: when the server cannot map the cache dir to a config key, registeredKey is null.
    // The delete button must not show (there is nothing safe to pass to the CLI).
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([
      makeMarketplace({ registeredKey: null, isOwn: false, isBuiltIn: false }),
    ]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    await screen.findByText('My Plugins');
    expect(screen.queryByRole('button', { name: /remove marketplace/i })).toBeNull();
  });

  it('delete button is HIDDEN for the OCA marketplace (isOwn = true)', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([
      makeMarketplace({
        name: 'office-coding-agent',
        registeredKey: 'office-coding-agent',
        isOwn: true,
        isBuiltIn: false,
      }),
    ]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    await screen.findByText('office-coding-agent');
    expect(screen.queryByRole('button', { name: /remove marketplace/i })).toBeNull();
  });

  it('delete button is HIDDEN for a built-in marketplace', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([
      makeMarketplace({
        name: 'awesome-copilot',
        slug: 'awesome-copilot',
        registeredKey: 'awesome-copilot',
        isBuiltIn: true,
        isOwn: false,
      }),
    ]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    await screen.findByText('awesome-copilot');
    expect(screen.queryByRole('button', { name: /remove marketplace/i })).toBeNull();
  });

  it('clicking delete calls removeMarketplace with the registeredKey', async () => {
    const user = userEvent.setup();
    const REGISTERED_KEY = 'stbrnner-microsoft/spt-iq';

    vi.mocked(pluginService.getMarketplaces)
      .mockResolvedValueOnce([
        makeMarketplace({ slug: 'stbrnner-microsoft-spt-iq', registeredKey: REGISTERED_KEY }),
      ])
      .mockResolvedValue([]); // after removal list is refreshed

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    const deleteBtn = await screen.findByRole('button', { name: /remove marketplace/i });
    await user.click(deleteBtn);

    expect(pluginService.removeMarketplace).toHaveBeenCalledOnce();
    expect(pluginService.removeMarketplace).toHaveBeenCalledWith(REGISTERED_KEY);
  });

  it('after successful removal the marketplace list is refreshed', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces)
      .mockResolvedValueOnce([makeMarketplace({ registeredKey: 'sbroenne/my-plugins' })])
      .mockResolvedValue([]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    const deleteBtn = await screen.findByRole('button', { name: /remove marketplace/i });
    await user.click(deleteBtn);

    // After the removal + refresh the item is gone
    await waitFor(() => {
      expect(screen.queryByText('My Plugins')).toBeNull();
    });
  });

  it('lock icon is shown for the OCA marketplace', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([
      makeMarketplace({ name: 'office-coding-agent', isOwn: true }),
    ]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    // The lock codicon is rendered as a <i class="codicon codicon-lock ..."> element
    const nameCell = await screen.findByText('office-coding-agent');
    // The parent row should contain the lock icon
    expect(nameCell.closest('div')?.querySelector('.codicon-lock')).toBeTruthy();
  });

  it('lock icon is shown for built-in marketplaces', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([
      makeMarketplace({ name: 'awesome-copilot', isBuiltIn: true, isOwn: false }),
    ]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    const nameCell = await screen.findByText('awesome-copilot');
    expect(nameCell.closest('div')?.querySelector('.codicon-lock')).toBeTruthy();
  });

  it('shows "No marketplaces registered" when list is empty', async () => {
    const user = userEvent.setup();
    vi.mocked(pluginService.getMarketplaces).mockResolvedValue([]);

    renderWithProviders(<PluginHub open onClose={() => {}} />);
    await openMarketplacesTab(user);

    await screen.findByText(/no marketplaces registered/i);
  });
});
