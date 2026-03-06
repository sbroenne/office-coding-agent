import React, { useCallback, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import { parseMcpJsonFile } from '@/services/mcp';
import { useMcpStatusStore } from '@/stores';
import { BUNDLED_MCP_SERVERS } from '@/types';
import type { McpServerConfig, McpServerStatus } from '@/types';
import { McpAddServerForm } from './McpAddServerForm';
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
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [showAddForm, setShowAddForm] = useState(false);
  const [editingServer, setEditingServer] = useState<string | null>(null);
  const [expandedTools, setExpandedTools] = useState<Set<string>>(new Set());
  const [logServer, setLogServer] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const mcpServers = useMcpStatusStore(s => s.servers);

  // Only bundled servers are shown — imported servers are managed via Plugin Hub
  const allServers: McpServerConfig[] = [...BUNDLED_MCP_SERVERS];
  const allNames = new Set(allServers.map(s => s.name));

  const getServerStatus = (name: string): McpServerStatus | 'disabled' => {
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

  const handleImportJson = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setImportStatus(null);
    setImportError(null);
    setIsImporting(true);

    try {
      const servers = await parseMcpJsonFile(file);
      setImportStatus(
        `Imported ${servers.length} server${servers.length === 1 ? '' : 's'} from ${file.name}.`
      );
    } catch (error) {
      setImportError(error instanceof Error ? error.message : 'Failed to import mcp.json.');
    } finally {
      setIsImporting(false);
      event.target.value = '';
    }
  }, []);

  const handleAddServer = (config: {
    name: string;
    description?: string;
    transport: 'http' | 'sse' | 'stdio';
    command?: string;
    args?: string[];
    url?: string;
    headers?: Record<string, string>;
  }) => {
    void config;
    setShowAddForm(false);
  };

  const handleEditServer = (config: {
    name: string;
    description?: string;
    transport: 'http' | 'sse' | 'stdio';
    command?: string;
    args?: string[];
    url?: string;
    headers?: Record<string, string>;
  }) => {
    void config;
    setEditingServer(null);
  };

  return (
    <div className="space-y-3 p-3">
      {/* Action bar */}
      <div className="flex items-center justify-between gap-2">
        <h4 className="text-xs font-medium text-muted-foreground">Servers ({allServers.length})</h4>
        <div className="flex items-center gap-1">
          <Button
            variant="secondary"
            size="sm"
            onClick={() => {
              setShowAddForm(true);
              setEditingServer(null);
            }}
          >
            <Codicon name="add" className="text-sm" />
            Add
          </Button>
          <>
            <input
              ref={inputRef}
              type="file"
              accept=".json,application/json"
              className="hidden"
              aria-label="Import mcp.json file"
              onChange={event => void handleImportJson(event)}
            />
            <Button
              variant="secondary"
              size="sm"
              onClick={() => inputRef.current?.click()}
              disabled={isImporting}
              aria-busy={isImporting}
            >
              {isImporting ? (
                <Codicon name="loading" className="text-sm codicon-modifier-spin" />
              ) : (
                <Codicon name="cloud-upload" className="text-sm" />
              )}
              {isImporting ? 'Importing…' : 'Import'}
            </Button>
          </>
        </div>
      </div>

      {/* Add server form */}
      {showAddForm && (
        <McpAddServerForm
          existingNames={allNames}
          onSubmit={handleAddServer}
          onCancel={() => setShowAddForm(false)}
        />
      )}

      {importStatus && (
        <div
          role="status"
          aria-live="polite"
          className="rounded-md border border-[var(--vscode-textLink-foreground)]/30 bg-[var(--vscode-textLink-foreground)]/10 px-3 py-2 text-xs text-[var(--vscode-textLink-foreground)]"
        >
          {importStatus}
        </div>
      )}
      {importError && (
        <div
          role="alert"
          aria-live="assertive"
          className="rounded-md border border-[var(--vscode-errorForeground)]/30 bg-[var(--vscode-errorForeground)]/10 px-3 py-2 text-xs text-[var(--vscode-errorForeground)]"
        >
          {importError}
        </div>
      )}

      {/* Server list */}
      <div className="space-y-1">
        {allServers.length === 0 ? (
          <p className="text-xs text-muted-foreground">
            No MCP servers configured. Add a server or import a <code>mcp.json</code> to get
            started.
          </p>
        ) : (
          allServers.map(server => {
            const status = getServerStatus(server.name);
            const serverState = mcpServers[server.name];
            const toolCount = serverState?.tools.length ?? 0;
            const isEditing = editingServer === server.name;
            const isExpanded = expandedTools.has(server.name);
            const bundled = isBundled(server.name);

            if (isEditing) {
              return (
                <McpAddServerForm
                  key={`edit-${server.name}`}
                  editMode
                  initial={server}
                  existingNames={allNames}
                  onSubmit={handleEditServer}
                  onCancel={() => setEditingServer(null)}
                />
              );
            }

            return (
              <div key={`mcp-server-${server.name}`} className="rounded-md border border-border">
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

      {/* Log viewer */}
      {logServer !== null && <McpLogViewer serverName={logServer} />}
    </div>
  );
};
