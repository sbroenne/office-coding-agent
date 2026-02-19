import React, { useCallback, useRef, useState } from 'react';
import { Loader2, ServerIcon, Trash2, Upload } from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { parseMcpJsonFile } from '@/services/mcp';
import { useSettingsStore } from '@/stores';

interface McpManagerDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export const McpManagerDialog: React.FC<McpManagerDialogProps> = ({ open, onOpenChange }) => {
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const importedMcpServers = useSettingsStore(s => s.importedMcpServers);
  const activeMcpServerNames = useSettingsStore(s => s.activeMcpServerNames);
  const importMcpServers = useSettingsStore(s => s.importMcpServers);
  const removeMcpServer = useSettingsStore(s => s.removeMcpServer);
  const toggleMcpServer = useSettingsStore(s => s.toggleMcpServer);

  const isServerActive = (name: string) =>
    activeMcpServerNames === null || activeMcpServerNames.includes(name);

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

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-[420px] max-h-[85vh] flex flex-col">
        <DialogHeader>
          <DialogTitle>MCP Servers</DialogTitle>
          <DialogDescription>
            Import a <code>mcp.json</code> file to connect AI tools from external MCP servers (HTTP
            or SSE transport).
          </DialogDescription>
        </DialogHeader>

        <div className="flex-1 overflow-y-auto space-y-3 pr-1">
          <div className="flex items-center justify-between gap-2">
            <h4 className="text-xs font-medium text-muted-foreground">Configured Servers</h4>
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
                {isImporting ? 'Importing…' : 'Import mcp.json'}
              </Button>
            </>
          </div>

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

          <div className="space-y-1">
            {importedMcpServers.length === 0 ? (
              <p className="text-xs text-muted-foreground">
                No MCP servers configured. Import a <code>mcp.json</code> to get started.
              </p>
            ) : (
              importedMcpServers.map(server => (
                <div
                  key={`mcp-server-${server.name}`}
                  className="flex items-center justify-between rounded-md border border-border px-2 py-1.5 gap-2"
                >
                  <button
                    className="flex items-center gap-2 min-w-0 flex-1 text-left"
                    onClick={() => toggleMcpServer(server.name)}
                    aria-pressed={isServerActive(server.name)}
                    title={isServerActive(server.name) ? 'Click to disable' : 'Click to enable'}
                  >
                    <ServerIcon
                      className={`size-3.5 shrink-0 ${isServerActive(server.name) ? 'text-emerald-500' : 'text-muted-foreground'}`}
                    />
                    <div className="min-w-0">
                      <p className="truncate text-sm font-medium">{server.name}</p>
                      <p className="truncate text-xs text-muted-foreground">
                        {server.description ?? server.url}
                      </p>
                    </div>
                  </button>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-7 px-2 text-destructive hover:text-destructive shrink-0"
                    onClick={() => removeMcpServer(server.name)}
                  >
                    <Trash2 className="size-3.5" />
                    Remove
                  </Button>
                </div>
              ))
            )}
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
};
