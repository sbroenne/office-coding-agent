import React, { useCallback, useRef, useState } from 'react';
import { Loader2, Trash2, Upload } from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { getBundledAgents } from '@/services/agents';
import { parseAgentsZipFile } from '@/services/extensions/zipImportService';
import { useSettingsStore } from '@/stores';

interface AgentManagerDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export const AgentManagerDialog: React.FC<AgentManagerDialogProps> = ({ open, onOpenChange }) => {
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const importedAgents = useSettingsStore(s => s.importedAgents);
  const importAgents = useSettingsStore(s => s.importAgents);
  const removeImportedAgent = useSettingsStore(s => s.removeImportedAgent);

  const bundledAgents = getBundledAgents();

  const handleImportZip = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;

      setImportStatus(null);
      setImportError(null);
      setIsImporting(true);

      try {
        const agents = await parseAgentsZipFile(file);
        importAgents(agents);
        setImportStatus(
          `Imported ${agents.length} agent${agents.length === 1 ? '' : 's'} from ${file.name}.`
        );
      } catch (error) {
        setImportError(error instanceof Error ? error.message : 'Failed to import agents ZIP.');
      } finally {
        setIsImporting(false);
        event.target.value = '';
      }
    },
    [importAgents]
  );

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-[420px] max-h-[85vh] flex flex-col">
        <DialogHeader>
          <DialogTitle>Manage Agents</DialogTitle>
          <DialogDescription>
            Import and manage custom agents. Bundled agents are shown separately and are read-only.
          </DialogDescription>
        </DialogHeader>

        <div className="flex-1 overflow-y-auto space-y-3 pr-1">
          <div className="flex items-center justify-between gap-2">
            <h4 className="text-xs font-medium text-muted-foreground">Custom Agents (ZIP)</h4>
            <>
              <input
                ref={inputRef}
                type="file"
                accept=".zip,application/zip"
                className="hidden"
                aria-label="Import agents ZIP file"
                onChange={event => void handleImportZip(event)}
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
                {isImporting ? 'Importing…' : 'Import Agents ZIP'}
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
            <p className="text-[11px] font-medium text-muted-foreground">Bundled (read-only)</p>
            {bundledAgents.length === 0 ? (
              <p className="text-xs text-muted-foreground">No bundled agents.</p>
            ) : (
              bundledAgents.map(agent => (
                <div
                  key={`bundled-agent-${agent.metadata.name}`}
                  className="flex items-center justify-between rounded-md border border-border px-2 py-1.5"
                >
                  <div className="min-w-0">
                    <p className="truncate text-sm font-medium">{agent.metadata.name}</p>
                    <p className="truncate text-xs text-muted-foreground">
                      {agent.metadata.description}
                    </p>
                  </div>
                  <span className="text-[10px] text-muted-foreground">Bundled</span>
                </div>
              ))
            )}
          </div>

          <div className="space-y-1">
            <p className="text-[11px] font-medium text-muted-foreground">Imported</p>
            {importedAgents.length === 0 ? (
              <p className="text-xs text-muted-foreground">No imported agents.</p>
            ) : (
              importedAgents.map(agent => (
                <div
                  key={`imported-agent-${agent.metadata.name}`}
                  className="flex items-center justify-between rounded-md border border-border px-2 py-1.5"
                >
                  <div className="min-w-0">
                    <p className="truncate text-sm font-medium">{agent.metadata.name}</p>
                    <p className="truncate text-xs text-muted-foreground">
                      {agent.metadata.description}
                    </p>
                  </div>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-7 px-2 text-destructive hover:text-destructive"
                    onClick={() => removeImportedAgent(agent.metadata.name)}
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
