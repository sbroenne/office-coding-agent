import React, { useState, useEffect, useCallback } from 'react';
import { Codicon } from '@/components/Codicon';
import type { InstalledPlugin, BrowsePlugin, RegisteredMarketplace } from '@/types/plugin';
import * as pluginService from '@/services/plugins/pluginService';
import PluginCard from './PluginCard';
import PluginDetailPanel from './PluginDetailPanel';

export interface PluginHubProps {
  open: boolean;
  onClose: () => void;
}

type TabId = 'installed' | 'browse' | 'marketplaces';

const TABS: { id: TabId; label: string }[] = [
  { id: 'installed', label: 'Installed' },
  { id: 'browse', label: 'Browse' },
  { id: 'marketplaces', label: 'Marketplaces' },
];

/** Collapsible section with a chevron toggle. */
const CollapsibleSection: React.FC<{
  title: string;
  count: number;
  defaultOpen?: boolean;
  children: React.ReactNode;
}> = ({ title, count, defaultOpen = true, children }) => {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div>
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex w-full items-center gap-1 px-1 py-1 text-[11px] font-medium text-muted-foreground transition-colors hover:bg-accent"
      >
        <Codicon name={open ? 'chevron-down' : 'chevron-right'} className="text-[12px]" />
        {title} ({count})
      </button>
      {open && children}
    </div>
  );
};

const PluginHub: React.FC<PluginHubProps> = ({ open, onClose }) => {
  const [activeTab, setActiveTab] = useState<TabId>('installed');
  const [installedPlugins, setInstalledPlugins] = useState<InstalledPlugin[]>([]);
  const [marketplaces, setMarketplaces] = useState<RegisteredMarketplace[]>([]);
  const [browsePlugins, setBrowsePlugins] = useState<BrowsePlugin[]>([]);
  const [selectedMarketplace, setSelectedMarketplace] = useState<string>('awesome-copilot');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedPlugin, setSelectedPlugin] = useState<string | null>(null);
  const [installSpec, setInstallSpec] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // ─── Data fetching─────────────────────────────────────────────

  const fetchInstalled = useCallback(async () => {
    try {
      const data = await pluginService.getInstalledPlugins();
      setInstalledPlugins(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load installed plugins');
    }
  }, []);

  const fetchMarketplaces = useCallback(async () => {
    try {
      const data = await pluginService.getMarketplaces();
      setMarketplaces(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load marketplaces');
    }
  }, []);

  const fetchBrowse = useCallback(
    async (marketplace: string) => {
      setLoading(true);
      setError(null);
      try {
        const data = await pluginService.browseMarketplace(marketplace);
        setBrowsePlugins(data);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to browse marketplace');
      } finally {
        setLoading(false);
      }
    },
    [],
  );

  const refreshAll = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      await Promise.all([fetchInstalled(), fetchMarketplaces()]);
    } finally {
      setLoading(false);
    }
  }, [fetchInstalled, fetchMarketplaces]);

  // On mount / when opened
  useEffect(() => {
    if (open) void refreshAll();
  }, [open, refreshAll]);

  // Browse tab: fetch when marketplace selection changes
  useEffect(() => {
    if (open && activeTab === 'browse' && selectedMarketplace) {
      void fetchBrowse(selectedMarketplace);
    }
  }, [open, activeTab, selectedMarketplace, fetchBrowse]);

  // ─── Actions ───────────────────────────────────────────────────

  const handleInstall = useCallback(
    async (spec: string) => {
      setLoading(true);
      setError(null);
      try {
        await pluginService.installPlugin(spec);
        await fetchInstalled();
        if (activeTab === 'browse') await fetchBrowse(selectedMarketplace);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Install failed');
      } finally {
        setLoading(false);
      }
    },
    [fetchInstalled, fetchBrowse, activeTab, selectedMarketplace],
  );

  const handleUninstall = useCallback(
    async (name: string) => {
      setLoading(true);
      setError(null);
      try {
        await pluginService.uninstallPlugin(name);
        await fetchInstalled();
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Uninstall failed');
      } finally {
        setLoading(false);
      }
    },
    [fetchInstalled],
  );

  const handleEnable = useCallback(
    async (name: string) => {
      try {
        await pluginService.enablePlugin(name);
        await fetchInstalled();
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Enable failed');
      }
    },
    [fetchInstalled],
  );

  const handleDisable = useCallback(
    async (name: string) => {
      try {
        await pluginService.disablePlugin(name);
        await fetchInstalled();
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Disable failed');
      }
    },
    [fetchInstalled],
  );

  const handleInlineInstall = useCallback(
    async (e: React.FormEvent) => {
      e.preventDefault();
      const spec = installSpec.trim();
      if (!spec) return;
      await handleInstall(spec);
      setInstallSpec('');
    },
    [installSpec, handleInstall],
  );

  const handleAddMarketplace = useCallback(
    async (e: React.FormEvent) => {
      e.preventDefault();
      const spec = installSpec.trim();
      if (!spec) return;
      setLoading(true);
      setError(null);
      try {
        await pluginService.addMarketplace(spec);
        await fetchMarketplaces();
        setInstallSpec('');
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to add marketplace');
      } finally {
        setLoading(false);
      }
    },
    [installSpec, fetchMarketplaces],
  );

  const handleRemoveMarketplace = useCallback(
    async (name: string) => {
      setLoading(true);
      setError(null);
      try {
        await pluginService.removeMarketplace(name);
        await fetchMarketplaces();
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to remove marketplace');
      } finally {
        setLoading(false);
      }
    },
    [fetchMarketplaces],
  );

  // ─── Filtering ─────────────────────────────────────────────────

  const query = searchQuery.toLowerCase();

  const filterPlugin = (p: { name: string; description?: string | null; manifest?: { description?: string } | null }) => {
    if (!query) return true;
    const name = p.name.toLowerCase();
    const desc = (
      ('manifest' in p && p.manifest?.description) ||
      ('description' in p && typeof p.description === 'string' && p.description) ||
      ''
    ).toLowerCase();
    return name.includes(query) || desc.includes(query);
  };

  const bundledPlugins = installedPlugins
    .filter((p) => p.marketplace === 'office-coding-agent')
    .filter(filterPlugin);
  const userPlugins = installedPlugins
    .filter((p) => p.marketplace !== 'office-coding-agent')
    .filter(filterPlugin);
  const filteredBrowse = browsePlugins.filter(filterPlugin);

  // ─── Detail view ───────────────────────────────────────────────

  if (!open) return null;

  if (selectedPlugin) {
    return (
      <div className="flex flex-col" style={{ height: '100%' }}>
        <PluginDetailPanel
          pluginName={selectedPlugin}
          onBack={() => setSelectedPlugin(null)}
          onActionComplete={() => void fetchInstalled()}
        />
      </div>
    );
  }

  // ─── Render ────────────────────────────────────────────────────

  return (
    <div
      className="flex flex-col"
      style={{
        height: '100%',
        background: 'var(--vscode-sideBar-background)',
        color: 'var(--vscode-sideBar-foreground)',
      }}
    >
      {/* ── Fixed header ─────────────────────────────────────── */}
      <div
        className="flex shrink-0 items-center gap-2 px-3"
        style={{
          height: 35,
          borderBottom: '1px solid var(--vscode-panel-border, var(--vscode-widget-border))',
        }}
      >
        <Codicon name="extensions" className="text-[14px]" />
        <span className="flex-1 text-[13px] font-semibold">Plugins</span>
        <button
          onClick={() => void refreshAll()}
          className="inline-flex h-[22px] w-[22px] items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
          style={{ color: 'var(--vscode-icon-foreground)' }}
          title="Refresh"
          disabled={loading}
        >
          <Codicon
            name="refresh"
            className={`text-[14px] ${loading ? 'codicon-modifier-spin' : ''}`}
          />
        </button>
        <button
          onClick={onClose}
          className="inline-flex h-[22px] w-[22px] items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
          style={{ color: 'var(--vscode-icon-foreground)' }}
          title="Close"
        >
          <Codicon name="close" className="text-[14px]" />
        </button>
      </div>

      {/* ── Search bar ───────────────────────────────────────── */}
      <div className="shrink-0 px-3 py-1.5">
        <div
          className="flex items-center gap-1.5 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1"
          style={{
            background: 'var(--vscode-input-background)',
            border: '1px solid var(--vscode-input-border, transparent)',
          }}
        >
          <Codicon
            name="search"
            className="shrink-0 text-[13px]"
            aria-hidden
          />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search plugins…"
            className="flex-1 bg-transparent text-[12px] outline-none"
            style={{
              color: 'var(--vscode-input-foreground)',
            }}
          />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery('')}
              className="shrink-0 text-muted-foreground hover:text-foreground"
              aria-label="Clear search"
            >
              <Codicon name="close" className="text-[12px]" />
            </button>
          )}
        </div>
      </div>

      {/* ── Tab bar ──────────────────────────────────────────── */}
      <div
        className="flex shrink-0 gap-0 px-3"
        style={{
          borderBottom: '1px solid var(--vscode-panel-border, var(--vscode-widget-border))',
        }}
      >
        {TABS.map((tab) => (
          <button
            key={tab.id}
            onClick={() => {
              setActiveTab(tab.id);
              setError(null);
            }}
            className="relative px-2 py-1 text-[11px] font-medium transition-colors"
            style={{
              color:
                activeTab === tab.id
                  ? 'var(--vscode-tab-activeForeground, var(--vscode-foreground))'
                  : 'var(--vscode-descriptionForeground)',
              background:
                activeTab === tab.id
                  ? 'var(--vscode-tab-activeBackground, transparent)'
                  : 'transparent',
            }}
          >
            {tab.label}
            {activeTab === tab.id && (
              <span
                className="absolute bottom-0 left-0 right-0 h-[2px]"
                style={{ background: 'var(--vscode-focusBorder)' }}
              />
            )}
          </button>
        ))}
      </div>

      {/* ── Error banner ─────────────────────────────────────── */}
      {error && (
        <div
          role="alert"
          className="mx-3 mt-2 rounded-md border border-[var(--vscode-errorForeground)]/30 bg-[var(--vscode-errorForeground)]/10 px-3 py-2 text-xs text-[var(--vscode-errorForeground)]"
        >
          {error}
        </div>
      )}

      {/* ── Scrollable content ───────────────────────────────── */}
      <div className="flex-1 overflow-y-auto">
        {/* ── Installed tab ──────────────────────────────────── */}
        {activeTab === 'installed' && (
          <div>
            {installedPlugins.length === 0 && !loading && (
              <p className="px-3 py-4 text-center text-xs text-muted-foreground">
                No plugins installed yet.
              </p>
            )}

            {bundledPlugins.length > 0 && (
              <CollapsibleSection
                title="Bundled"
                count={bundledPlugins.length}
                defaultOpen
              >
                {bundledPlugins.map((p) => (
                  <PluginCard
                    key={p.name}
                    plugin={p}
                    mode="installed"
                    onEnable={(name) => void handleEnable(name)}
                    onDisable={(name) => void handleDisable(name)}
                    onUninstall={(name) => void handleUninstall(name)}
                    onClick={(name) => setSelectedPlugin(name)}
                  />
                ))}
              </CollapsibleSection>
            )}

            {userPlugins.length > 0 && (
              <CollapsibleSection
                title="Installed"
                count={userPlugins.length}
                defaultOpen
              >
                {userPlugins.map((p) => (
                  <PluginCard
                    key={p.name}
                    plugin={p}
                    mode="installed"
                    onEnable={(name) => void handleEnable(name)}
                    onDisable={(name) => void handleDisable(name)}
                    onUninstall={(name) => void handleUninstall(name)}
                    onClick={(name) => setSelectedPlugin(name)}
                  />
                ))}
              </CollapsibleSection>
            )}
          </div>
        )}

        {/* ── Browse tab ─────────────────────────────────────── */}
        {activeTab === 'browse' && (
          <div>
            {/* Marketplace selector */}
            <div className="px-3 py-2">
              <select
                value={selectedMarketplace}
                onChange={(e) => setSelectedMarketplace(e.target.value)}
                className="w-full rounded-[var(--vscode-cornerRadius-small)] px-2 py-1 text-[12px] outline-none"
                style={{
                  background: 'var(--vscode-input-background)',
                  color: 'var(--vscode-input-foreground)',
                  border: '1px solid var(--vscode-input-border, transparent)',
                }}
              >
                {marketplaces.map((m) => (
                  <option key={m.name} value={m.name}>
                    {m.name}
                    {m.pluginCount != null ? ` (${m.pluginCount})` : ''}
                  </option>
                ))}
              </select>
            </div>

            {loading && (
              <div className="flex items-center justify-center gap-2 py-4">
                <Codicon name="loading" className="text-[14px] codicon-modifier-spin" />
                <span className="text-xs text-muted-foreground">Loading plugins…</span>
              </div>
            )}

            {!loading && filteredBrowse.length === 0 && (
              <p className="px-3 py-4 text-center text-xs text-muted-foreground">
                {searchQuery
                  ? 'No plugins match your search.'
                  : 'No plugins found in this marketplace.'}
              </p>
            )}

            {!loading &&
              filteredBrowse.map((p) => (
                <PluginCard
                  key={p.name}
                  plugin={p}
                  mode="browse"
                  onInstall={(spec) => void handleInstall(spec)}
                  onClick={(name) => setSelectedPlugin(name)}
                />
              ))}
          </div>
        )}

        {/* ── Marketplaces tab ───────────────────────────────── */}
        {activeTab === 'marketplaces' && (
          <div className="space-y-1 px-3 py-2">
            {marketplaces.length === 0 && !loading && (
              <p className="py-4 text-center text-xs text-muted-foreground">
                No marketplaces registered.
              </p>
            )}

            {marketplaces.map((m) => (
              <div
                key={m.name}
                className="flex items-center justify-between gap-2 rounded-md border px-2 py-1.5"
                style={{
                  borderColor: 'var(--vscode-panel-border, var(--vscode-widget-border))',
                }}
              >
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-1.5">
                    <span className="truncate text-[12px] font-medium">{m.name}</span>
                    {m.isBuiltIn && (
                      <Codicon
                        name="lock"
                        className="shrink-0 text-[11px] text-muted-foreground"
                        aria-hidden
                      />
                    )}
                  </div>
                  <p className="truncate text-[10px] text-muted-foreground">
                    {m.source}
                    {m.pluginCount != null ? ` · ${m.pluginCount} plugins` : ''}
                  </p>
                </div>
                {!m.isBuiltIn && (
                  <button
                    onClick={() => void handleRemoveMarketplace(m.name)}
                    className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-[var(--vscode-errorForeground)]"
                    title="Remove marketplace"
                    disabled={loading}
                  >
                    <Codicon name="trash" className="text-[12px]" />
                  </button>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ── Fixed footer ─────────────────────────────────────── */}
      {activeTab === 'installed' && (
        <form
          onSubmit={(e) => void handleInlineInstall(e)}
          className="flex shrink-0 items-center gap-1.5 px-3 py-2"
          style={{
            borderTop: '1px solid var(--vscode-panel-border, var(--vscode-widget-border))',
          }}
        >
          <div
            className="flex flex-1 items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1"
            style={{
              background: 'var(--vscode-input-background)',
              border: '1px solid var(--vscode-input-border, transparent)',
            }}
          >
            <Codicon name="add" className="shrink-0 text-[12px] text-muted-foreground" />
            <input
              type="text"
              value={installSpec}
              onChange={(e) => setInstallSpec(e.target.value)}
              placeholder="owner/repo, name@marketplace, or local path"
              className="flex-1 bg-transparent text-[12px] outline-none"
              style={{ color: 'var(--vscode-input-foreground)' }}
            />
          </div>
          <button
            type="submit"
            disabled={!installSpec.trim() || loading}
            className="shrink-0 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1 text-[11px] font-medium transition-colors disabled:opacity-50"
            style={{
              background: 'var(--vscode-button-background)',
              color: 'var(--vscode-button-foreground)',
            }}
            onMouseEnter={(e) => {
              (e.currentTarget as HTMLElement).style.background =
                'var(--vscode-button-hoverBackground, var(--vscode-button-background))';
            }}
            onMouseLeave={(e) => {
              (e.currentTarget as HTMLElement).style.background =
                'var(--vscode-button-background)';
            }}
          >
            {loading ? (
              <Codicon name="loading" className="text-[11px] codicon-modifier-spin" />
            ) : (
              'Install'
            )}
          </button>
        </form>
      )}

      {activeTab === 'marketplaces' && (
        <form
          onSubmit={(e) => void handleAddMarketplace(e)}
          className="flex shrink-0 items-center gap-1.5 px-3 py-2"
          style={{
            borderTop: '1px solid var(--vscode-panel-border, var(--vscode-widget-border))',
          }}
        >
          <div
            className="flex flex-1 items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1"
            style={{
              background: 'var(--vscode-input-background)',
              border: '1px solid var(--vscode-input-border, transparent)',
            }}
          >
            <Codicon name="add" className="shrink-0 text-[12px] text-muted-foreground" />
            <input
              type="text"
              value={installSpec}
              onChange={(e) => setInstallSpec(e.target.value)}
              placeholder="owner/repo or local path"
              className="flex-1 bg-transparent text-[12px] outline-none"
              style={{ color: 'var(--vscode-input-foreground)' }}
            />
          </div>
          <button
            type="submit"
            disabled={!installSpec.trim() || loading}
            className="shrink-0 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1 text-[11px] font-medium transition-colors disabled:opacity-50"
            style={{
              background: 'var(--vscode-button-background)',
              color: 'var(--vscode-button-foreground)',
            }}
            onMouseEnter={(e) => {
              (e.currentTarget as HTMLElement).style.background =
                'var(--vscode-button-hoverBackground, var(--vscode-button-background))';
            }}
            onMouseLeave={(e) => {
              (e.currentTarget as HTMLElement).style.background =
                'var(--vscode-button-background)';
            }}
          >
            {loading ? (
              <Codicon name="loading" className="text-[11px] codicon-modifier-spin" />
            ) : (
              'Add'
            )}
          </button>
        </form>
      )}
    </div>
  );
};

export default PluginHub;
