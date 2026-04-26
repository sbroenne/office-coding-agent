import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import type { BrowsePlugin, InstalledPlugin, PluginMarketplaceSummary } from '@/types/plugin';
import {
  addMarketplace,
  browseMarketplace,
  getInstalledPlugins,
  getMarketplaces,
  installPlugin,
  removeMarketplace,
  uninstallPlugin,
  updateAllPlugins,
  updateMarketplace,
  updatePlugin,
} from '@/services/plugins/pluginService';

type PluginView = 'installed' | 'browse';

const iconButtonClass =
  'inline-flex h-6 w-6 items-center justify-center rounded-[var(--vscode-cornerRadius-small)] text-[var(--vscode-icon-foreground)] transition-colors hover:bg-accent focus-visible:outline focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)]';

const actionButtonClass =
  'inline-flex h-6 items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] border border-border px-2 text-xs transition-colors hover:bg-accent focus-visible:outline focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)]';

export const PluginManagerPanel: React.FC = () => {
  const [view, setView] = useState<PluginView>('installed');
  const [installed, setInstalled] = useState<InstalledPlugin[]>([]);
  const [marketplaces, setMarketplaces] = useState<PluginMarketplaceSummary[]>([]);
  const [selectedMarketplace, setSelectedMarketplace] = useState('');
  const [browseItems, setBrowseItems] = useState<BrowsePlugin[]>([]);
  const [installSpec, setInstallSpec] = useState('');
  const [marketplaceSpec, setMarketplaceSpec] = useState('');
  const [filter, setFilter] = useState('');
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const refreshInstalled = useCallback(async () => {
    const plugins = await getInstalledPlugins();
    setInstalled(plugins);
  }, []);

  const refreshMarketplaces = useCallback(async () => {
    const next = await getMarketplaces();
    setMarketplaces(next);
    setSelectedMarketplace(current => current || next[0]?.name || '');
  }, []);

  const refreshAll = useCallback(async () => {
    await Promise.all([refreshInstalled(), refreshMarketplaces()]);
  }, [refreshInstalled, refreshMarketplaces]);

  useEffect(() => {
    void refreshAll().catch(error => {
      setStatus(error instanceof Error ? error.message : String(error));
    });
  }, [refreshAll]);

  useEffect(() => {
    if (view !== 'browse' || !selectedMarketplace) return;
    void browseMarketplace(selectedMarketplace)
      .then(setBrowseItems)
      .catch(error => {
        setStatus(error instanceof Error ? error.message : String(error));
        setBrowseItems([]);
      });
  }, [selectedMarketplace, view]);

  const filteredInstalled = useMemo(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return installed;
    return installed.filter(
      plugin =>
        plugin.name.toLowerCase().includes(q) ||
        plugin.marketplace?.toLowerCase().includes(q) ||
        plugin.manifest?.description?.toLowerCase().includes(q)
    );
  }, [filter, installed]);

  const filteredBrowse = useMemo(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return browseItems;
    return browseItems.filter(
      plugin =>
        plugin.name.toLowerCase().includes(q) ||
        (plugin.description?.toLowerCase().includes(q) ?? false) ||
        plugin.marketplace.toLowerCase().includes(q)
    );
  }, [browseItems, filter]);

  const runAction = async (
    label: string,
    action: () => Promise<{ success: boolean; message: string }>
  ) => {
    setBusy(true);
    setStatus(`${label}…`);
    try {
      const result = await action();
      setStatus(result.message);
      await refreshAll();
      if (view === 'browse' && selectedMarketplace) {
        setBrowseItems(await browseMarketplace(selectedMarketplace));
      }
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="flex h-full flex-col">
      <div className="border-b border-border p-2">
        <div className="mb-2 flex items-center gap-1">
          <button
            className={`${actionButtonClass} ${view === 'installed' ? 'bg-accent' : ''}`}
            onClick={() => setView('installed')}
          >
            Installed
          </button>
          <button
            className={`${actionButtonClass} ${view === 'browse' ? 'bg-accent' : ''}`}
            onClick={() => setView('browse')}
          >
            Browse
          </button>
          <button
            className={iconButtonClass}
            onClick={() => void refreshAll()}
            disabled={busy}
            aria-label="Refresh plugins"
            title="Refresh plugins"
          >
            <Codicon name="refresh" className="text-[13px]" />
          </button>
        </div>

        <input
          value={filter}
          onChange={event => setFilter(event.target.value)}
          placeholder="Search plugins"
          className="h-7 w-full rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-input px-2 text-sm outline-none focus:border-[var(--vscode-focusBorder)]"
        />
      </div>

      {status && (
        <div className="border-b border-border px-3 py-1.5 text-xs text-muted-foreground">
          {status}
        </div>
      )}

      {view === 'installed' ? (
        <InstalledPluginsView
          plugins={filteredInstalled}
          installSpec={installSpec}
          setInstallSpec={setInstallSpec}
          busy={busy}
          onInstall={() => {
            void runAction('Installing plugin', async () => {
              const result = await installPlugin(installSpec);
              setInstallSpec('');
              return result;
            });
          }}
          onUninstall={plugin => {
            void runAction('Uninstalling plugin', () =>
              uninstallPlugin(plugin.name, plugin.marketplace)
            );
          }}
          onUpdate={plugin => {
            void runAction('Updating plugin', () => updatePlugin(plugin.name, plugin.marketplace));
          }}
          onUpdateAll={() => {
            void runAction('Updating plugins', updateAllPlugins);
          }}
        />
      ) : (
        <BrowsePluginsView
          marketplaces={marketplaces}
          selectedMarketplace={selectedMarketplace}
          setSelectedMarketplace={setSelectedMarketplace}
          items={filteredBrowse}
          marketplaceSpec={marketplaceSpec}
          setMarketplaceSpec={setMarketplaceSpec}
          busy={busy}
          onInstall={plugin => {
            void runAction('Installing plugin', () =>
              installPlugin(`${plugin.name}@${plugin.marketplace}`)
            );
          }}
          onAddMarketplace={() => {
            void runAction('Adding marketplace', async () => {
              const result = await addMarketplace(marketplaceSpec);
              setMarketplaceSpec('');
              return result;
            });
          }}
          onRemoveMarketplace={marketplace => {
            void runAction('Removing marketplace', () =>
              removeMarketplace(marketplace.slug, marketplace.registeredKey)
            );
          }}
          onUpdateMarketplace={marketplace => {
            void runAction('Updating marketplace', () => updateMarketplace(marketplace.name));
          }}
        />
      )}
    </div>
  );
};

const InstalledPluginsView: React.FC<{
  plugins: InstalledPlugin[];
  installSpec: string;
  setInstallSpec: (value: string) => void;
  busy: boolean;
  onInstall: () => void;
  onUninstall: (plugin: InstalledPlugin) => void;
  onUpdate: (plugin: InstalledPlugin) => void;
  onUpdateAll: () => void;
}> = ({
  plugins,
  installSpec,
  setInstallSpec,
  busy,
  onInstall,
  onUninstall,
  onUpdate,
  onUpdateAll,
}) => (
  <div className="flex-1 overflow-y-auto p-2">
    <div className="mb-2 flex gap-1">
      <input
        value={installSpec}
        onChange={event => setInstallSpec(event.target.value)}
        placeholder="owner/repo or plugin@marketplace"
        className="h-7 min-w-0 flex-1 rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-input px-2 text-sm outline-none focus:border-[var(--vscode-focusBorder)]"
      />
      <button
        className={actionButtonClass}
        onClick={onInstall}
        disabled={busy || !installSpec.trim()}
      >
        <Codicon name="add" className="text-[13px]" />
        Install
      </button>
      <button
        className={actionButtonClass}
        onClick={onUpdateAll}
        disabled={busy || plugins.length === 0}
      >
        Update all
      </button>
    </div>

    {plugins.length === 0 ? (
      <p className="p-2 text-xs text-muted-foreground">No plugins installed.</p>
    ) : (
      <div className="space-y-1">
        {plugins.map(plugin => (
          <PluginCard
            key={`${plugin.name}-${plugin.marketplace}`}
            title={plugin.name}
            subtitle={plugin.manifest?.description ?? plugin.marketplace}
            badges={[
              plugin.marketplace || 'local',
              `${plugin.components?.agentCount ?? 0} agents`,
              `${plugin.components?.skillCount ?? 0} skills`,
              `${plugin.components?.commandCount ?? 0} prompts`,
              `${plugin.components?.mcpServerCount ?? 0} MCP`,
            ]}
            actions={
              <>
                <button
                  className={iconButtonClass}
                  onClick={() => onUpdate(plugin)}
                  disabled={busy}
                  title="Update"
                >
                  <Codicon name="cloud-download" className="text-[13px]" />
                </button>
                <button
                  className={iconButtonClass}
                  onClick={() => onUninstall(plugin)}
                  disabled={busy}
                  title="Uninstall"
                >
                  <Codicon name="trash" className="text-[13px]" />
                </button>
              </>
            }
          />
        ))}
      </div>
    )}
  </div>
);

const BrowsePluginsView: React.FC<{
  marketplaces: PluginMarketplaceSummary[];
  selectedMarketplace: string;
  setSelectedMarketplace: (value: string) => void;
  items: BrowsePlugin[];
  marketplaceSpec: string;
  setMarketplaceSpec: (value: string) => void;
  busy: boolean;
  onInstall: (plugin: BrowsePlugin) => void;
  onAddMarketplace: () => void;
  onRemoveMarketplace: (marketplace: PluginMarketplaceSummary) => void;
  onUpdateMarketplace: (marketplace: PluginMarketplaceSummary) => void;
}> = ({
  marketplaces,
  selectedMarketplace,
  setSelectedMarketplace,
  items,
  marketplaceSpec,
  setMarketplaceSpec,
  busy,
  onInstall,
  onAddMarketplace,
  onRemoveMarketplace,
  onUpdateMarketplace,
}) => (
  <div className="flex-1 overflow-y-auto p-2">
    <div className="mb-2 flex gap-1">
      <input
        value={marketplaceSpec}
        onChange={event => setMarketplaceSpec(event.target.value)}
        placeholder="marketplace source"
        className="h-7 min-w-0 flex-1 rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-input px-2 text-sm outline-none focus:border-[var(--vscode-focusBorder)]"
      />
      <button
        className={actionButtonClass}
        onClick={onAddMarketplace}
        disabled={busy || !marketplaceSpec.trim()}
      >
        <Codicon name="add" className="text-[13px]" />
        Add
      </button>
    </div>

    <div className="mb-2 flex gap-1">
      <select
        value={selectedMarketplace}
        onChange={event => setSelectedMarketplace(event.target.value)}
        className="h-7 min-w-0 flex-1 rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-input px-2 text-sm outline-none focus:border-[var(--vscode-focusBorder)]"
      >
        {marketplaces.map(marketplace => (
          <option key={marketplace.name} value={marketplace.name}>
            {marketplace.name}
          </option>
        ))}
      </select>
      {marketplaces
        .filter(marketplace => marketplace.name === selectedMarketplace)
        .map(marketplace => (
          <React.Fragment key={marketplace.name}>
            <button
              className={iconButtonClass}
              onClick={() => onUpdateMarketplace(marketplace)}
              disabled={busy}
              title="Update marketplace"
            >
              <Codicon name="sync" className="text-[13px]" />
            </button>
            {!marketplace.isBuiltIn && !marketplace.isOwn && (
              <button
                className={iconButtonClass}
                onClick={() => onRemoveMarketplace(marketplace)}
                disabled={busy}
                title="Remove marketplace"
              >
                <Codicon name="trash" className="text-[13px]" />
              </button>
            )}
          </React.Fragment>
        ))}
    </div>

    {items.length === 0 ? (
      <p className="p-2 text-xs text-muted-foreground">No marketplace plugins found.</p>
    ) : (
      <div className="space-y-1">
        {items.map(plugin => (
          <PluginCard
            key={`${plugin.name}-${plugin.marketplace}`}
            title={plugin.name}
            subtitle={plugin.description}
            badges={[plugin.marketplace, plugin.installed ? 'installed' : 'available']}
            actions={
              <button
                className={actionButtonClass}
                onClick={() => onInstall(plugin)}
                disabled={busy || plugin.installed}
              >
                <Codicon
                  name={plugin.installed ? 'check' : 'cloud-download'}
                  className="text-[13px]"
                />
                {plugin.installed ? 'Installed' : 'Install'}
              </button>
            }
          />
        ))}
      </div>
    )}
  </div>
);

const PluginCard: React.FC<{
  title: string;
  subtitle?: string;
  badges: string[];
  actions: React.ReactNode;
}> = ({ title, subtitle, badges, actions }) => (
  <div className="rounded-[var(--vscode-cornerRadius-medium)] border border-border p-2">
    <div className="flex items-start gap-2">
      <Codicon
        name="extensions"
        className="mt-0.5 shrink-0 text-[14px] text-[var(--vscode-icon-foreground)]"
      />
      <div className="min-w-0 flex-1">
        <div className="truncate text-sm font-medium">{title}</div>
        {subtitle && <div className="truncate text-xs text-muted-foreground">{subtitle}</div>}
        <div className="mt-1 flex flex-wrap gap-1">
          {badges.filter(Boolean).map(badge => (
            <span
              key={badge}
              className="rounded-[var(--vscode-cornerRadius-small)] bg-[var(--vscode-badge-background)] px-1.5 py-0.5 text-[10px] text-[var(--vscode-badge-foreground)]"
            >
              {badge}
            </span>
          ))}
        </div>
      </div>
      <div className="flex shrink-0 items-center gap-1">{actions}</div>
    </div>
  </div>
);
