import React, { useCallback } from 'react';
import { cn } from '@/lib/utils';
import { Codicon } from '@/components/Codicon';
import type { InstalledPlugin, BrowsePlugin } from '@/types/plugin';

function isInstalledPlugin(p: InstalledPlugin | BrowsePlugin): p is InstalledPlugin {
  return 'enabled' in p && 'cachePath' in p;
}

function getSourceLabel(plugin: InstalledPlugin | BrowsePlugin): string {
  if (isInstalledPlugin(plugin)) {
    if (plugin.marketplace === 'office-coding-agent') return 'bundled';
    if (!plugin.marketplace) return 'direct';
    return plugin.marketplace;
  }
  return plugin.marketplace || 'marketplace';
}

function getComponentSummary(plugin: InstalledPlugin): string {
  const parts: string[] = [];
  const { agentCount, skillCount } = plugin.components;
  if (agentCount) parts.push(`${agentCount} agent${agentCount !== 1 ? 's' : ''}`);
  if (skillCount) parts.push(`${skillCount} skill${skillCount !== 1 ? 's' : ''}`);
  return parts.join(', ');
}

export interface PluginCardProps {
  plugin: InstalledPlugin | BrowsePlugin;
  mode: 'installed' | 'browse';
  onInstall?: (spec: string) => void;
  onUninstall?: (name: string) => void;
  onEnable?: (name: string) => void;
  onDisable?: (name: string) => void;
  onClick?: (name: string) => void;
}

const PluginCard: React.FC<PluginCardProps> = ({
  plugin,
  mode,
  onInstall,
  onEnable,
  onDisable,
  onClick,
}) => {
  const installed = isInstalledPlugin(plugin);
  const name = plugin.name;
  const version = installed ? plugin.version : plugin.version;
  const description = installed
    ? plugin.manifest?.description ?? ''
    : plugin.description ?? '';
  const source = getSourceLabel(plugin);

  const handleCardClick = useCallback(() => onClick?.(name), [onClick, name]);

  const handleAction = useCallback(
    (e: React.MouseEvent) => {
      e.stopPropagation();
      if (mode === 'browse' && !installed) {
        const spec = plugin.marketplace ? `${name}@${plugin.marketplace}` : name;
        onInstall?.(spec);
      } else if (installed && plugin.enabled) {
        onDisable?.(name);
      } else if (installed && !plugin.enabled) {
        onEnable?.(name);
      }
    },
    [mode, installed, plugin, name, onInstall, onDisable, onEnable],
  );

  return (
    <div
      role="button"
      tabIndex={0}
      className={cn('flex cursor-pointer items-start gap-3 px-3 py-2 transition-colors')}
      style={{
        borderBottom: '1px solid var(--vscode-panel-border, var(--vscode-widget-border))',
      }}
      onClick={handleCardClick}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          handleCardClick();
        }
      }}
      onMouseEnter={(e) => {
        (e.currentTarget as HTMLElement).style.background =
          'var(--vscode-list-hoverBackground)';
      }}
      onMouseLeave={(e) => {
        (e.currentTarget as HTMLElement).style.background = 'transparent';
      }}
    >
      {/* Icon */}
      <Codicon
        name="package"
        className="mt-0.5 shrink-0 text-[16px]"
      />

      {/* Body */}
      <div className="flex min-w-0 flex-1 flex-col gap-0">
        {/* Name */}
        <span
          className="truncate text-[13px] font-semibold leading-[18px]"
          style={{ color: 'var(--vscode-foreground)' }}
        >
          {name}
        </span>

        {/* Meta line */}
        <span
          className="truncate text-[11px] leading-[16px]"
          style={{ color: 'var(--vscode-descriptionForeground)' }}
        >
          {version ? `v${version}` : 'v0.0.0'} · {source}
        </span>

        {/* Description */}
        {description && (
          <span
            className="mt-0.5 truncate text-[12px] leading-[16px]"
            style={{ color: 'var(--vscode-descriptionForeground)' }}
          >
            {description}
          </span>
        )}

        {/* Component summary (installed only) */}
        {installed && (
          <span
            className="mt-0.5 text-[11px] leading-[14px]"
            style={{ color: 'var(--vscode-descriptionForeground)', opacity: 0.7 }}
          >
            {getComponentSummary(plugin as InstalledPlugin)}
          </span>
        )}
      </div>

      {/* Action button */}
      {mode === 'browse' && !(plugin as BrowsePlugin).installed && (
        <button
          type="button"
          className="shrink-0 self-center rounded-[var(--vscode-cornerRadius-small)] px-2 py-0.5 text-[11px] font-medium transition-colors"
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
          onClick={handleAction}
        >
          Install
        </button>
      )}

      {installed && plugin.enabled && (
        <button
          type="button"
          className="shrink-0 self-center rounded-[var(--vscode-cornerRadius-small)] p-1 transition-colors hover:bg-accent"
          style={{ color: 'var(--vscode-testing-iconPassed, #28a745)' }}
          title="Disable plugin"
          onClick={handleAction}
        >
          <Codicon name="check" className="text-[14px]" />
        </button>
      )}

      {installed && !plugin.enabled && (
        <button
          type="button"
          className="shrink-0 self-center rounded-[var(--vscode-cornerRadius-small)] p-1 transition-colors hover:bg-accent"
          style={{ color: 'var(--vscode-descriptionForeground)' }}
          title="Enable plugin"
          onClick={handleAction}
        >
          <Codicon name="circle-slash" className="text-[14px]" />
        </button>
      )}
    </div>
  );
};

export default PluginCard;
