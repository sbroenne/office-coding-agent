import React, { useCallback, useRef, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  Loader2,
  Pencil,
  Plus,
  RefreshCw,
  Square,
  Play,
  Trash2,
  Upload,
  FileText,
} from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { parseMcpJsonFile } from '@/services/mcp';
import { useSettingsStore, useMcpStatusStore } from '@/stores';
import { BUNDLED_MCP_SERVERS } from '@/types';
import type { McpServerConfig, McpServerStatus } from '@/types';
import { McpAddServerForm } from './McpAddServerForm';
import { McpLogViewer } from './McpLogViewer';

interface McpManagerDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

const STATUS_COLORS: Record<McpServerStatus | 'disabled', string> = {
  connected: 'bg-emerald-500',
  starting: 'bg-amber-400',
  error: 'bg-red-500',
  stopped: 'bg-zinc-400',
  disabled: 'bg-zinc-300 dark:bg-zinc-600',
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

export const McpManagerDialog: React.FC<McpManagerDialogProps> = ({ open, onOpenChange }) => {
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [showAddForm, setShowAddForm] = useState(false);
  const [editingServer, setEditingServer] = useState<string | null>(null);
  const [expandedTools, setExpandedTools] = useState<Set<string>>(new Set());
  const [logServer, setLogServer] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const importedMcpServers = useSettingsStore(s => s.importedMcpServers);
  const activeMcpServerNames = useSettingsStore(s => s.activeMcpServerNames);
  const importMcpServers = useSettingsStore(s => s.importMcpServers);
  const removeMcpServer = useSettingsStore(s => s.removeMcpServer);
  const toggleMcpServer = useSettingsStore(s => s.toggleMcpServer);
  const updateMcpServer = useSettingsStore(s => s.updateMcpServer);
  const mcpServers = useMcpStatusStore(s => s.servers);

  // Merge bundled + imported for display
  const bundledNames = new Set(BUNDLED_MCP_SERVERS.map(s => s.name));
  const allServers: McpServerConfig[] = [
    ...BUNDLED_MCP_SERVERS,
    ...importedMcpServers.filter(s => !bundledNames.has(s.name)),
  ];
  const allNames = new Set(allServers.map(s => s.name));

  const isServerActive = (name: string) =>
    activeMcpServerNames === null || activeMcpServerNames.includes(name);

  const getServerStatus = (name: string): McpServerStatus | 'disabled' => {
    if (!isServerActive(name)) return 'disabled';
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

  const handleImportJson = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;

      setImportStatus(null);
      setImportError(null);
      setIsImporting(true);

      try {
        const servers = await parseMcpJsonFile(file);
        importMcpServers(servers);
        setImportStatus(
          `Imported ${servers.length} server${servers.length === 1 ? '' : 's'} from ${file.name}.`
        );
      } catch (error) {
        setImportError(error instanceof Error ? error.message : 'Failed to import mcp.json.');
      } finally {
        setIsImporting(false);
        event.target.value = '';
      }
    },
    [importMcpServers]
  );

  const handleAddServer = (config: {
    name: string;
    description?: string;
    transport: 'http' | 'sse' | 'stdio';
    command?: string;
    args?: string[];
    url?: string;
    headers?: Record<string, string>;
  }) => {
    importMcpServers([config as McpServerConfig]);
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
    updateMcpServer(config.name, config as Partial<McpServerConfig>);
    setEditingServer(null);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-[480px] max-h-[85vh] flex flex-col">
        <DialogHeader>
          <DialogTitle>MCP Servers</DialogTitle>
          <DialogDescription>
            Manage MCP servers that provide additional AI tools. Import from <code>mcp.json</code>{' '}
            or add manually.
          </DialogDescription>
        </DialogHeader>

        <div className="flex-1 overflow-y-auto space-y-3 pr-1">
          {/* Action bar */}
          <div className="flex items-center justify-between gap-2">
            <h4 className="text-xs font-medium text-muted-foreground">
              Servers ({allServers.length})
            </h4>
            <div className="flex items-center gap-1">
              <Button
                variant="secondary"
                size="sm"
                onClick={() => {
                  setShowAddForm(true);
                  setEditingServer(null);
                }}
              >
                <Plus className="size-3.5" />
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
                    <Loader2 className="size-3.5 animate-spin" />
                  ) : (
                    <Upload className="size-3.5" />
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
              className="rounded-md border border-emerald-300 bg-emerald-50 px-3 py-2 text-xs text-emerald-900 dark:border-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-100"
            >
              {importStatus}
            </div>
          )}
          {importError && (
            <div
              role="alert"
              aria-live="assertive"
              className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-xs text-red-900 dark:border-red-700 dark:bg-red-900/30 dark:text-red-100"
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
                  <div
                    key={`mcp-server-${server.name}`}
                    className="rounded-md border border-border"
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
                              <span className="shrink-0 rounded-full bg-blue-100 px-1.5 py-0 text-[9px] font-medium text-blue-700 dark:bg-blue-900/40 dark:text-blue-300">
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
                            <p className="truncate text-[10px] text-red-500">{serverState.error}</p>
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
                              <ChevronDown className="size-3" />
                            ) : (
                              <ChevronRight className="size-3" />
                            )}
                            {toolCount}
                          </button>
                        )}
                        {/* Show output */}
                        <button
                          onClick={() =>
                            setLogServer(logServer === server.name ? null : server.name)
                          }
                          className={`inline-flex h-6 w-6 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-accent-foreground ${logServer === server.name ? 'bg-accent text-accent-foreground' : ''}`}
                          title="Show Output"
                        >
                          <FileText className="size-3" />
                        </button>
                        {/* Restart (re-toggle) */}
                        {isServerActive(server.name) && (
                          <button
                            onClick={() => {
                              toggleMcpServer(server.name);
                              setTimeout(() => toggleMcpServer(server.name), 100);
                            }}
                            className="inline-flex h-6 w-6 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-accent-foreground"
                            title="Restart"
                          >
                            <RefreshCw className="size-3" />
                          </button>
                        )}
                        {/* Start/Stop */}
                        <button
                          onClick={() => toggleMcpServer(server.name)}
                          className="inline-flex h-6 w-6 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-accent-foreground"
                          title={isServerActive(server.name) ? 'Stop' : 'Start'}
                        >
                          {isServerActive(server.name) ? (
                            <Square className="size-3" />
                          ) : (
                            <Play className="size-3" />
                          )}
                        </button>
                        {/* Edit (imported only) */}
                        {!bundled && (
                          <button
                            onClick={() => {
                              setEditingServer(server.name);
                              setShowAddForm(false);
                            }}
                            className="inline-flex h-6 w-6 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-accent-foreground"
                            title="Edit"
                          >
                            <Pencil className="size-3" />
                          </button>
                        )}
                        {/* Remove (imported only) */}
                        {!bundled && (
                          <button
                            onClick={() => removeMcpServer(server.name)}
                            className="inline-flex h-6 w-6 items-center justify-center rounded text-destructive hover:bg-accent"
                            title="Remove"
                          >
                            <Trash2 className="size-3" />
                          </button>
                        )}
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
      </DialogContent>
    </Dialog>
  );
};
