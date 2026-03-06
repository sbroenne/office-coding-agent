import React, { useState, useEffect, useCallback } from 'react';
import { Codicon } from '@/components/Codicon';
import * as pluginService from '@/services/plugins/pluginService';
import type { InstalledPlugin, PluginManifest, PluginComponents } from '@/types/plugin';

export interface PluginDetailPanelProps {
  pluginName: string;
  onBack: () => void;
  onActionComplete?: () => void;
}

interface PluginDetails {
  plugin: InstalledPlugin;
  manifest: PluginManifest | null;
  components: PluginComponents;
}

const PluginDetailPanel: React.FC<PluginDetailPanelProps> = ({
  pluginName,
  onBack,
  onActionComplete,
}) => {
  const [details, setDetails] = useState<PluginDetails | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [actionLoading, setActionLoading] = useState<string | null>(null);

  const fetchDetails = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await pluginService.getPluginDetails(pluginName);
      setDetails(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load plugin details');
    } finally {
      setLoading(false);
    }
  }, [pluginName]);

  useEffect(() => {
    void fetchDetails();
  }, [fetchDetails]);

  const runAction = useCallback(
    async (action: string, fn: () => Promise<unknown>) => {
      setActionLoading(action);
      try {
        await fn();
        onActionComplete?.();
        // Refresh details after toggle actions; go back after uninstall
        if (action === 'uninstall') {
          onBack();
        } else {
          await fetchDetails();
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : `Failed to ${action} plugin`);
      } finally {
        setActionLoading(null);
      }
    },
    [onActionComplete, onBack, fetchDetails],
  );

  const handleToggle = useCallback(() => {
    if (!details) return;
    const { plugin } = details;
    if (plugin.enabled) {
      void runAction('disable', () => pluginService.disablePlugin(plugin.name));
    } else {
      void runAction('enable', () => pluginService.enablePlugin(plugin.name));
    }
  }, [details, runAction]);

  const handleUpdate = useCallback(() => {
    void runAction('update', () => pluginService.updatePlugin(pluginName));
  }, [pluginName, runAction]);

  const handleUninstall = useCallback(() => {
    void runAction('uninstall', () => pluginService.uninstallPlugin(pluginName));
  }, [pluginName, runAction]);

  if (loading) {
    return (
      <div className="space-y-3 p-3">
        <button
          onClick={onBack}
          className="inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1 py-0.5 text-[12px] transition-colors hover:bg-accent"
          style={{ color: 'var(--vscode-textLink-foreground)' }}
        >
          <Codicon name="arrow-left" className="text-[12px]" />
          Back
        </button>
        <div className="flex items-center gap-2 py-6 justify-center">
          <Codicon name="loading" className="text-[14px] codicon-modifier-spin" />
          <span className="text-xs text-muted-foreground">Loading plugin details…</span>
        </div>
      </div>
    );
  }

  if (error && !details) {
    return (
      <div className="space-y-3 p-3">
        <button
          onClick={onBack}
          className="inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1 py-0.5 text-[12px] transition-colors hover:bg-accent"
          style={{ color: 'var(--vscode-textLink-foreground)' }}
        >
          <Codicon name="arrow-left" className="text-[12px]" />
          Back
        </button>
        <div
          role="alert"
          className="rounded-md border border-[var(--vscode-errorForeground)]/30 bg-[var(--vscode-errorForeground)]/10 px-3 py-2 text-xs text-[var(--vscode-errorForeground)]"
        >
          {error}
        </div>
      </div>
    );
  }

  if (!details) return null;

  const { plugin, manifest, components } = details;
  const isBundled = plugin.marketplace === 'office-coding-agent';
  const description = manifest?.description ?? '';
  const author = manifest?.author;
  const keywords = manifest?.keywords ?? [];
  const hasComponents =
    components.agentCount > 0 || components.skillCount > 0 || components.mcpServerCount > 0;

  return (
    <div className="space-y-3 p-3">
      {/* Back button */}
      <button
        onClick={onBack}
        className="inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1 py-0.5 text-[12px] transition-colors hover:bg-accent"
        style={{ color: 'var(--vscode-textLink-foreground)' }}
      >
        <Codicon name="arrow-left" className="text-[12px]" />
        Back
      </button>

      {/* Header: name + toggle */}
      <div className="flex items-start justify-between gap-2">
        <div className="flex items-center gap-2 min-w-0">
          <Codicon name="package" className="shrink-0 text-[18px]" />
          <div className="min-w-0">
            <h3 className="truncate text-[13px] font-semibold">{plugin.name}</h3>
            {isBundled && (
              <span className="rounded-full bg-[var(--vscode-textLink-foreground)]/15 px-1.5 py-0 text-[9px] font-medium text-[var(--vscode-textLink-foreground)]">
                Bundled
              </span>
            )}
          </div>
        </div>
        <button
          onClick={handleToggle}
          disabled={actionLoading !== null}
          className="shrink-0 rounded-[var(--vscode-cornerRadius-small)] px-2 py-0.5 text-[11px] font-medium transition-colors"
          style={{
            background: plugin.enabled
              ? 'var(--vscode-button-secondaryBackground, var(--vscode-input-background))'
              : 'var(--vscode-button-background)',
            color: plugin.enabled
              ? 'var(--vscode-button-secondaryForeground, var(--vscode-foreground))'
              : 'var(--vscode-button-foreground)',
          }}
        >
          {actionLoading === 'enable' || actionLoading === 'disable' ? (
            <Codicon name="loading" className="text-[11px] codicon-modifier-spin" />
          ) : plugin.enabled ? (
            'Disable'
          ) : (
            'Enable'
          )}
        </button>
      </div>

      {/* Error banner */}
      {error && (
        <div
          role="alert"
          className="rounded-md border border-[var(--vscode-errorForeground)]/30 bg-[var(--vscode-errorForeground)]/10 px-3 py-2 text-xs text-[var(--vscode-errorForeground)]"
        >
          {error}
        </div>
      )}

      {/* Metadata */}
      <div
        className="space-y-1 rounded-md border px-3 py-2 text-[11px]"
        style={{
          borderColor: 'var(--vscode-panel-border, var(--vscode-widget-border))',
          color: 'var(--vscode-descriptionForeground)',
        }}
      >
        <div className="flex justify-between">
          <span>Version</span>
          <span style={{ color: 'var(--vscode-foreground)' }}>{plugin.version || '0.0.0'}</span>
        </div>
        {author && (
          <div className="flex justify-between">
            <span>Author</span>
            <span style={{ color: 'var(--vscode-foreground)' }}>{author.name}</span>
          </div>
        )}
        <div className="flex justify-between">
          <span>Source</span>
          <span style={{ color: 'var(--vscode-foreground)' }}>
            {plugin.marketplace || 'direct install'}
          </span>
        </div>
      </div>

      {/* Description */}
      {description && (
        <div>
          <h4 className="mb-1 text-[11px] font-medium text-muted-foreground">Description</h4>
          <p className="text-[12px] leading-[16px]" style={{ color: 'var(--vscode-foreground)' }}>
            {description}
          </p>
        </div>
      )}

      {/* Components */}
      {hasComponents && (
        <div>
          <h4 className="mb-1 text-[11px] font-medium text-muted-foreground">Components</h4>
          <div
            className="space-y-1 rounded-md border px-3 py-2 text-[11px]"
            style={{ borderColor: 'var(--vscode-panel-border, var(--vscode-widget-border))' }}
          >
            {components.agentCount > 0 && (
              <div>
                <span className="font-medium" style={{ color: 'var(--vscode-foreground)' }}>
                  Agents
                </span>
                <span className="text-muted-foreground">
                  {' '}
                  — {components.agentNames.join(', ')}
                </span>
              </div>
            )}
            {components.skillCount > 0 && (
              <div>
                <span className="font-medium" style={{ color: 'var(--vscode-foreground)' }}>
                  Skills
                </span>
                <span className="text-muted-foreground">
                  {' '}
                  — {components.skillNames.join(', ')}
                </span>
              </div>
            )}
            {components.mcpServerCount > 0 && (
              <div>
                <span className="font-medium" style={{ color: 'var(--vscode-foreground)' }}>
                  MCP Servers
                </span>
                <span className="text-muted-foreground">
                  {' '}
                  — {components.mcpServerNames.join(', ')}
                </span>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Keywords */}
      {keywords.length > 0 && (
        <div>
          <h4 className="mb-1 text-[11px] font-medium text-muted-foreground">Keywords</h4>
          <div className="flex flex-wrap gap-1">
            {keywords.map((kw) => (
              <span
                key={kw}
                className="rounded-full px-1.5 py-0 text-[10px]"
                style={{
                  background: 'var(--vscode-badge-background)',
                  color: 'var(--vscode-badge-foreground)',
                }}
              >
                {kw}
              </span>
            ))}
          </div>
        </div>
      )}

      {/* Action buttons */}
      {!isBundled && (
        <div
          className="flex gap-2 border-t pt-3"
          style={{ borderColor: 'var(--vscode-panel-border, var(--vscode-widget-border))' }}
        >
          <button
            onClick={handleUpdate}
            disabled={actionLoading !== null}
            className="flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1 text-[11px] font-medium transition-colors"
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
            {actionLoading === 'update' ? (
              <Codicon name="loading" className="text-[11px] codicon-modifier-spin" />
            ) : (
              <Codicon name="refresh" className="text-[11px]" />
            )}
            Update
          </button>
          <button
            onClick={handleUninstall}
            disabled={actionLoading !== null}
            className="flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-2 py-1 text-[11px] font-medium transition-colors"
            style={{
              background: 'var(--vscode-errorForeground)',
              color: 'var(--vscode-editor-background, #fff)',
            }}
          >
            {actionLoading === 'uninstall' ? (
              <Codicon name="loading" className="text-[11px] codicon-modifier-spin" />
            ) : (
              <Codicon name="trash" className="text-[11px]" />
            )}
            Uninstall
          </button>
        </div>
      )}
    </div>
  );
};

export default PluginDetailPanel;
