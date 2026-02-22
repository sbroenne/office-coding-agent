import React, { useCallback, useRef, useState } from 'react';
import { Download, Loader2, Trash2, Upload } from 'lucide-react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { getBundledAgents } from '@/services/agents';
import { parseAgentsZipFile, parseAgentMarkdownFile } from '@/services/extensions/zipImportService';
import { downloadAgent, downloadAgentsZip } from '@/services/extensions/zipExportService';
import { useSettingsStore } from '@/stores';

interface AgentManagerDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export const AgentManagerDialog: React.FC<AgentManagerDialogProps> = ({ open, onOpenChange }) => {
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [isDownloadingAll, setIsDownloadingAll] = useState(false);
  const zipInputRef = useRef<HTMLInputElement>(null);
  const mdInputRef = useRef<HTMLInputElement>(null);

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

  const handleImportMd = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (!file) return;

      setImportStatus(null);
      setImportError(null);
      setIsImporting(true);

      try {
        const agent = await parseAgentMarkdownFile(file);
        importAgents([agent]);
        setImportStatus(`Imported agent "${agent.metadata.name}" from ${file.name}.`);
      } catch (error) {
        setImportError(error instanceof Error ? error.message : 'Failed to import agent file.');
      } finally {
        setIsImporting(false);
        event.target.value = '';
      }
    },
    [importAgents]
  );

  const handleDownloadAll = useCallback(async () => {
    if (importedAgents.length === 0) return;
    setIsDownloadingAll(true);
    try {
      await downloadAgentsZip(importedAgents);
    } finally {
      setIsDownloadingAll(false);
    }
  }, [importedAgents]);

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
          {/* Import toolbar */}
          <div className="flex items-center justify-between gap-2">
            <h4 className="text-xs font-medium text-muted-foreground">Custom Agents</h4>
            <div className="flex items-center gap-1">
              <input
                ref={zipInputRef}
                type="file"
                accept=".zip,application/zip"
                className="hidden"
                aria-label="Import agents ZIP file"
                onChange={event => void handleImportZip(event)}
              />
              <input
                ref={mdInputRef}
                type="file"
                accept=".md,text/markdown"
                className="hidden"
                aria-label="Import agent Markdown file"
                onChange={event => void handleImportMd(event)}
              />
              <Button
                variant="secondary"
                size="sm"
                onClick={() => zipInputRef.current?.click()}
                disabled={isImporting}
                aria-busy={isImporting}
                title="Import agents from ZIP"
              >
                {isImporting ? (
                  <Loader2 className="size-3.5 animate-spin" />
                ) : (
                  <Upload className="size-3.5" />
                )}
                ZIP
              </Button>
              <Button
                variant="secondary"
                size="sm"
                onClick={() => mdInputRef.current?.click()}
                disabled={isImporting}
                aria-busy={isImporting}
                title="Import a single agent .md file"
              >
                <Upload className="size-3.5" />
                .md
              </Button>
            </div>
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

          {/* Bundled agents */}
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
                  <Button
                    variant="ghost"
                    size="icon"
                    className="size-7 shrink-0"
                    onClick={() => downloadAgent(agent)}
                    aria-label={`Download ${agent.metadata.name} as template`}
                    title="Download as template"
                  >
                    <Download className="size-3.5" />
                  </Button>
                </div>
              ))
            )}
          </div>

          {/* Imported agents */}
          <div className="space-y-1">
            <div className="flex items-center justify-between">
              <p className="text-[11px] font-medium text-muted-foreground">Imported</p>
              {importedAgents.length > 0 && (
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-6 gap-1 px-1.5 text-[11px]"
                  onClick={() => void handleDownloadAll()}
                  disabled={isDownloadingAll}
                  title="Download all custom agents as ZIP"
                >
                  {isDownloadingAll ? (
                    <Loader2 className="size-3 animate-spin" />
                  ) : (
                    <Download className="size-3" />
                  )}
                  Download all
                </Button>
              )}
            </div>
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
                  <div className="flex items-center gap-0.5 shrink-0">
                    <Button
                      variant="ghost"
                      size="icon"
                      className="size-7"
                      onClick={() => downloadAgent(agent)}
                      aria-label={`Download ${agent.metadata.name}`}
                      title="Download as .md"
                    >
                      <Download className="size-3.5" />
                    </Button>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="size-7 text-destructive hover:text-destructive"
                      onClick={() => removeImportedAgent(agent.metadata.name)}
                      aria-label={`Remove ${agent.metadata.name}`}
                      title="Remove"
                    >
                      <Trash2 className="size-3.5" />
                    </Button>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
};
