import React, { useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { useSettingsStore } from '@/stores/settingsStore';
import { useMcpStatusStore } from '@/stores';
import { BUNDLED_MCP_SERVERS } from '@/types';
import type { McpServerConfig, McpServerStatus } from '@/types';
import { McpLogViewer } from './McpLogViewer';

const STATUS_COLORS: Record<McpServerStatus | 'disabled', string> = {
  connected: 'bg-[var(--vscode-textLink-foreground)]',
  starting: 'bg-[var(--vscode-descriptionForeground)]',
  error: 'bg-[var(--vscode-errorForeground)]',
  stopped: 'bg-[var(--vscode-descriptionForeground)]',
  disabled: 'bg-[var(--vscode-descriptionForeground)] opacity-50',
};

const STATUS_LABELS: Record<McpServerStatus | 'disabled', string> = {
  connected: 'Connected',
  starting: 'Starting…',
  error: 'Error',
  stopped: 'Stopped',
  disabled: 'Disabled',
};

function isBundled(name: string): boolean {
  return BUNDLED_MCP_SERVERS.some(s => s.name === name);
}

export const McpManagerPanel: React.FC = () => {
  const [expandedTools, setExpandedTools] = useState<Set<string>>(new Set());
  const [logServer, setLogServer] = useState<string | null>(null);

  const mcpServers = useMcpStatusStore(s => s.servers);
  const toggleMcpServer = useSettingsStore(s => s.toggleMcpServer);
  const disabledMcpServerNames = useSettingsStore(s => s.disabledMcpServerNames);

  const allServers: McpServerConfig[] = [...BUNDLED_MCP_SERVERS];

  const getServerStatus = (name: string): McpServerStatus | 'disabled' => {
    if (disabledMcpServerNames.includes(name)) return 'disabled';
    return mcpServers[name]?.status ?? 'stopped';
  };

  const toggleTools = (name: string) => {
    setExpandedTools(prev => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  };

  return (
    <div className="space-y-3 p-3">
      {/* Header */}
      <div className="flex items-center justify-between gap-2">
        <h4 className="text-xs font-medium text-muted-foreground">Servers ({allServers.length})</h4>
      </div>

      {/* Server list */}
      <div className="space-y-1">
        {allServers.length === 0 ? (
          <p className="text-xs text-muted-foreground">No MCP servers configured.</p>
        ) : (
          allServers.map(server => {
            const status = getServerStatus(server.name);
            const serverState = mcpServers[server.name];
            const toolCount = serverState?.tools.length ?? 0;
            const isExpanded = expandedTools.has(server.name);
            const bundled = isBundled(server.name);

            return (
              <div
                key={`mcp-server-${server.name}`}
                className={`rounded-md border border-border transition-opacity ${disabledMcpServerNames.includes(server.name) ? 'opacity-50' : ''}`}
              >
                {/* Server row */}
                <div className="flex items-center justify-between px-2 py-1.5 gap-2">
                  <div className="flex items-center gap-2 min-w-0 flex-1">
                    {/* Status dot */}
                    <span
                      className={`size-2 shrink-0 rounded-full ${STATUS_COLORS[status]}`}
                      title={STATUS_LABELS[status]}
                    />
                    {/* Name + badge */}
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-1.5">
                        <span className="truncate text-sm font-medium">{server.name}</span>
                        {bundled && (
                          <span className="shrink-0 rounded-full bg-[var(--vscode-textLink-foreground)]/15 px-1.5 py-0 text-[9px] font-medium text-[var(--vscode-textLink-foreground)]">
                            Built-in
                          </span>
                        )}
                      </div>
                      <p className="truncate text-[10px] text-muted-foreground">
                        {server.description ??
                          (server.transport === 'stdio'
                            ? [server.command, ...(server.args ?? [])].join(' ')
                            : server.url)}
                      </p>
                      {status === 'error' && serverState?.error && (
                        <p className="truncate text-[10px] text-[var(--vscode-errorForeground)]">
                          {serverState.error}
                        </p>
                      )}
                    </div>
                  </div>

                  {/* Action buttons */}
                  <div className="flex items-center gap-0.5 shrink-0">
                    {/* Enable/disable toggle */}
                    <button
                      onClick={() => toggleMcpServer(server.name)}
                      className={`inline-flex h-6 w-6 items-center justify-center rounded transition-colors hover:bg-accent ${disabledMcpServerNames.includes(server.name) ? 'text-muted-foreground/40' : 'text-[var(--vscode-textLink-foreground)]'}`}
                      title={
                        disabledMcpServerNames.includes(server.name)
                          ? 'Enable server'
                          : 'Disable server'
                      }
                      aria-pressed={!disabledMcpServerNames.includes(server.name)}
                      aria-label={`Toggle ${server.name}`}
                    >
                      <Codicon
                        name={
                          disabledMcpServerNames.includes(server.name) ? 'circle-slash' : 'check'
                        }
                        className="text-xs"
                      />
                    </button>
                    {/* Tools toggle */}
                    {toolCount > 0 && (
                      <button
                        onClick={() => toggleTools(server.name)}
                        className="inline-flex h-6 items-center gap-0.5 rounded px-1 text-[10px] text-muted-foreground hover:bg-accent hover:text-accent-foreground"
                        title={`${toolCount} tool${toolCount === 1 ? '' : 's'}`}
                      >
                        {isExpanded ? (
                          <Codicon name="chevron-down" className="text-xs" />
                        ) : (
                          <Codicon name="chevron-right" className="text-xs" />
                        )}
                        {toolCount}
                      </button>
                    )}
                    {/* Show output */}
                    <button
                      onClick={() => setLogServer(logServer === server.name ? null : server.name)}
                      className={`inline-flex h-6 w-6 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-accent-foreground ${logServer === server.name ? 'bg-accent text-accent-foreground' : ''}`}
                      title="Show Output"
                    >
                      <Codicon name="output" className="text-xs" />
                    </button>
                  </div>
                </div>

                {/* Expanded tools list */}
                {isExpanded && serverState && serverState.tools.length > 0 && (
                  <div className="border-t border-border bg-muted/30 px-3 py-1.5">
                    <p className="mb-1 text-[10px] font-medium text-muted-foreground">
                      Tools ({serverState.tools.length})
                    </p>
                    <div className="space-y-0.5">
                      {serverState.tools.map(tool => (
                        <div key={tool.name} className="text-[10px]">
                          <span className="font-medium">{tool.name}</span>
                          {tool.description && (
                            <span className="text-muted-foreground"> — {tool.description}</span>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      {/* Plugin servers note */}
      <p className="text-xs text-muted-foreground">
        Additional MCP servers are configured via plugins.
      </p>

      {/* Log viewer */}
      {logServer !== null && <McpLogViewer serverName={logServer} />}
    </div>
  );
};
