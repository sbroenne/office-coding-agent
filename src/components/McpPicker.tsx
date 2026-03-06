/**
 * VS Code-style MCP server tools picker.
 *
 * Shows ALL available MCP servers (bundled + from installed plugins) with
 * per-server enable/disable toggles. Mirrors the "Select tools" quick-pick
 * in VS Code Copilot Chat.
 */
import React, { useState, useEffect } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores/settingsStore';
import { getLocalApiBase } from '@/lib/api';
import type { McpServerConfig } from '@/types/mcp';

interface McpPickerProps {
  onOpenPanel?: (panel: string) => void;
}

async function fetchAllMcpServers(): Promise<McpServerConfig[]> {
  try {
    const res = await fetch(`${getLocalApiBase()}/api/mcp-servers`);
    if (!res.ok) return [];
    const data = (await res.json()) as { servers: McpServerConfig[] };
    return data.servers ?? [];
  } catch {
    return [];
  }
}

export const McpPicker: React.FC<McpPickerProps> = ({ onOpenPanel }) => {
  const [open, setOpen] = useState(false);
  const [servers, setServers] = useState<McpServerConfig[]>([]);
  const toggleMcpServer = useSettingsStore(s => s.toggleMcpServer);
  const isMcpServerEnabled = useSettingsStore(s => s.isMcpServerEnabled);
  // Subscribe to disabledMcpServerNames so the component re-renders when toggles change
  useSettingsStore(s => s.disabledMcpServerNames);

  useEffect(() => {
    void fetchAllMcpServers().then(setServers);
  }, []);

  // Re-fetch when popover opens so it reflects newly installed plugins
  useEffect(() => {
    if (open) void fetchAllMcpServers().then(setServers);
  }, [open]);

  const enabledCount = servers.filter(s => isMcpServerEnabled(s.name)).length;
  const total = servers.length;
  const showBadge = total > 0 && enabledCount < total;

  return (
    <Popover.Root open={open} onOpenChange={setOpen}>
      <Popover.Trigger asChild>
        <button
          className="relative inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
          style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
          aria-label="MCP servers"
          title="MCP servers"
        >
          <Codicon name="server" className="text-[14px]" />
          {showBadge && (
            <span
              className="absolute -right-0.5 -top-0.5 flex h-3.5 min-w-3.5 items-center justify-center rounded-full bg-[var(--vscode-badge-background)] px-0.5 text-[8px] font-bold leading-none text-[var(--vscode-badge-foreground)]"
              aria-label={`${enabledCount} of ${total} MCP servers enabled`}
            >
              {enabledCount}
            </span>
          )}
        </button>
      </Popover.Trigger>

      <Popover.Portal>
        <Popover.Content
          className="z-50 w-72 max-h-80 overflow-y-auto rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
          sideOffset={4}
          align="end"
        >
          <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">MCP Servers</div>

          {servers.length === 0 ? (
            <div className="px-2 py-2 text-xs text-muted-foreground">
              No MCP servers available. Install a plugin to add MCP servers.
            </div>
          ) : (
            servers.map(server => {
              const enabled = isMcpServerEnabled(server.name);
              return (
                <button
                  key={server.name}
                  onClick={() => toggleMcpServer(server.name)}
                  className={cn(
                    'flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
                    !enabled && 'opacity-50'
                  )}
                  aria-pressed={enabled}
                  aria-label={server.name}
                  title={enabled ? `Disable ${server.name}` : `Enable ${server.name}`}
                >
                  <div
                    className={cn(
                      'mt-0.5 flex size-4 shrink-0 items-center justify-center rounded border',
                      enabled
                        ? 'border-[var(--vscode-textLink-foreground)] bg-[var(--vscode-textLink-foreground)]/10'
                        : 'border-border'
                    )}
                  >
                    {enabled && (
                      <Codicon
                        name="check"
                        className="text-xs text-[var(--vscode-textLink-foreground)]"
                      />
                    )}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="font-medium text-foreground">{server.name}</div>
                    {server.description && (
                      <div className="truncate text-xs text-muted-foreground">
                        {server.description}
                      </div>
                    )}
                    <div className="text-xs text-muted-foreground opacity-70">
                      {server.transport === 'stdio' ? server.command : server.url}
                    </div>
                  </div>
                </button>
              );
            })
          )}

          <div className="mt-1 border-t border-border pt-1">
            <button
              onClick={() => {
                setOpen(false);
                onOpenPanel?.('plugins');
              }}
              className="flex w-full items-center gap-1.5 rounded-md px-2 py-1.5 text-left text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
            >
              <Codicon name="settings-gear" className="text-[11px]" />
              <span>Manage MCP servers…</span>
            </button>
          </div>
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
