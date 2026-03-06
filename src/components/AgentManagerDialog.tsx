import React, { useCallback, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import { getBundledAgents } from '@/services/agents';
import { parseAgentsZipFile, parseAgentMarkdownFile } from '@/services/extensions/zipImportService';
import { downloadAgent } from '@/services/extensions/zipExportService';

export const AgentManagerPanel: React.FC = () => {
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const zipInputRef = useRef<HTMLInputElement>(null);
  const mdInputRef = useRef<HTMLInputElement>(null);

  const bundledAgents = getBundledAgents();

  const handleImportZip = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setImportStatus(null);
    setImportError(null);
    setIsImporting(true);

    try {
      const agents = await parseAgentsZipFile(file);
      setImportStatus(
        `Imported ${agents.length} agent${agents.length === 1 ? '' : 's'} from ${file.name}.`
      );
    } catch (error) {
      setImportError(error instanceof Error ? error.message : 'Failed to import agents ZIP.');
    } finally {
      setIsImporting(false);
      event.target.value = '';
    }
  }, []);

  const handleImportMd = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setImportStatus(null);
    setImportError(null);
    setIsImporting(true);

    try {
      const agent = await parseAgentMarkdownFile(file);
      setImportStatus(`Imported agent "${agent.metadata.name}" from ${file.name}.`);
    } catch (error) {
      setImportError(error instanceof Error ? error.message : 'Failed to import agent file.');
    } finally {
      setIsImporting(false);
      event.target.value = '';
    }
  }, []);

  return (
    <div className="space-y-3 p-3">
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
              <Codicon name="loading" className="text-sm codicon-modifier-spin" />
            ) : (
              <Codicon name="cloud-upload" className="text-sm" />
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
            <Codicon name="cloud-upload" className="text-sm" />
            .md
          </Button>
        </div>
      </div>

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
                <Codicon name="cloud-download" className="text-sm" />
              </Button>
            </div>
          ))
        )}
      </div>

      {/* Imported agents — managed via Plugin Hub */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">Imported</p>
        <p className="text-xs text-muted-foreground">No imported agents.</p>
      </div>
    </div>
  );
};
