/**
 * VS Code-style MCP server tools picker.
 *
 * Shows built-in MCP servers with per-server enable/disable toggles.
 * Mirrors the "Select tools" quick-pick in VS Code Copilot Chat.
 */
import React, { useState, useEffect } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores/settingsStore';
import { useMcpStatusStore } from '@/stores/mcpStatusStore';
import { getLocalApiBase } from '@/lib/api';
import type { McpServerConfig, McpServerState } from '@/types/mcp';
import { toMcpServerKey } from '@/utils/mcpServerKey';

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

interface McpPickerProps {
  onInitiateOAuth?: (serverName: string) => Promise<void>;
}

function getRuntimeState(
  states: Record<string, McpServerState>,
  serverName: string
): McpServerState | undefined {
  const serverKey = toMcpServerKey(serverName);
  return (
    states[serverName] ??
    states[serverKey] ??
    Object.entries(states).find(([name]) => toMcpServerKey(name) === serverKey)?.[1]
  );
}

function needsOAuthAction(server: McpServerConfig, state?: McpServerState): boolean {
  if (server.transport !== 'http' && server.transport !== 'sse') return false;
  if (server.headers?.Authorization) return false;
  return state?.status === 'needs-auth' || state?.status === 'failed' || state?.status === 'error';
}

export const McpPicker: React.FC<McpPickerProps> = ({ onInitiateOAuth }) => {
  const [open, setOpen] = useState(false);
  const [servers, setServers] = useState<McpServerConfig[]>([]);
  const [signingInServer, setSigningInServer] = useState<string | null>(null);
  const [authErrors, setAuthErrors] = useState<Record<string, string>>({});
  const toggleMcpServer = useSettingsStore(s => s.toggleMcpServer);
  const isMcpServerEnabled = useSettingsStore(s => s.isMcpServerEnabled);
  const runtimeStates = useMcpStatusStore(s => s.servers);
  // Subscribe to disabledMcpServerNames so the component re-renders when toggles change
  useSettingsStore(s => s.disabledMcpServerNames);

  useEffect(() => {
    void fetchAllMcpServers().then(setServers);
  }, []);

  // Re-fetch when popover opens so the built-in server list stays current.
  useEffect(() => {
    if (open) void fetchAllMcpServers().then(setServers);
  }, [open]);

  const enabledCount = servers.filter(s => isMcpServerEnabled(s.name)).length;
  const total = servers.length;
  const showBadge = total > 0 && enabledCount < total;
  const handleOAuthClick = (serverName: string) => {
    if (!onInitiateOAuth) return;
    setAuthErrors(errors => {
      const next = { ...errors };
      delete next[serverName];
      return next;
    });
    setSigningInServer(serverName);
    void onInitiateOAuth(serverName)
      .catch(error => {
        setAuthErrors(errors => ({
          ...errors,
          [serverName]: error instanceof Error ? error.message : String(error),
        }));
      })
      .finally(() => {
        setSigningInServer(null);
      });
  };

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
            <div className="px-2 py-2 text-xs text-muted-foreground space-y-1">
              <div>No MCP servers available.</div>
              <div>
                Manage Copilot CLI plugins from your terminal with{' '}
                <code className="font-mono">copilot plugin</code>.
              </div>
            </div>
          ) : (
            servers.map(server => {
              const enabled = isMcpServerEnabled(server.name);
              const state = getRuntimeState(runtimeStates, server.name);
              const showOAuthAction = enabled && onInitiateOAuth && needsOAuthAction(server, state);
              const isSigningIn = signingInServer === server.name;
              const authError = authErrors[server.name];
              return (
                <div
                  key={server.name}
                  className={cn(
                    'flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
                    !enabled && 'opacity-50'
                  )}
                >
                  <button
                    type="button"
                    onClick={() => toggleMcpServer(server.name)}
                    className="mt-0.5 flex size-4 shrink-0 items-center justify-center rounded border"
                    style={{
                      borderColor: enabled ? 'var(--vscode-textLink-foreground)' : undefined,
                      backgroundColor: enabled
                        ? 'color-mix(in srgb, var(--vscode-textLink-foreground) 10%, transparent)'
                        : undefined,
                    }}
                    aria-pressed={enabled}
                    aria-label={server.name}
                    title={enabled ? `Disable ${server.name}` : `Enable ${server.name}`}
                  >
                    {enabled && (
                      <Codicon
                        name="check"
                        className="text-xs text-[var(--vscode-textLink-foreground)]"
                      />
                    )}
                  </button>
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
                    {state?.status && state.status !== 'stopped' && (
                      <div className="text-xs text-muted-foreground">
                        Status: {state.status}
                        {state.error ? ` — ${state.error}` : ''}
                      </div>
                    )}
                    {authError && signingInServer === null && state?.status !== 'connected' && (
                      <div className="text-xs text-[var(--vscode-errorForeground)]">
                        {authError}
                      </div>
                    )}
                  </div>
                  {showOAuthAction && (
                    <button
                      type="button"
                      className="shrink-0 rounded px-2 py-0.5 text-xs text-[var(--vscode-button-foreground)]"
                      style={{ backgroundColor: 'var(--vscode-button-background)' }}
                      disabled={isSigningIn}
                      onClick={event => {
                        event.stopPropagation();
                        handleOAuthClick(server.name);
                      }}
                    >
                      {isSigningIn
                        ? 'Signing in...'
                        : state?.status === 'failed' || state?.status === 'error'
                          ? 'Retry sign in'
                          : 'Sign in'}
                    </button>
                  )}
                </div>
              );
            })
          )}
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
