import React, { useCallback, useRef, useState } from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import { getBundledSkills } from '@/services/skills';
import { parseSkillsZipFile, parseSkillMarkdownFile } from '@/services/extensions/zipImportService';
import { downloadSkill } from '@/services/extensions/zipExportService';

export const SkillManagerPanel: React.FC = () => {
  const [importStatus, setImportStatus] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const zipInputRef = useRef<HTMLInputElement>(null);
  const mdInputRef = useRef<HTMLInputElement>(null);

  const bundledSkills = getBundledSkills();

  const handleImportZip = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setImportStatus(null);
    setImportError(null);
    setIsImporting(true);

    try {
      const skills = await parseSkillsZipFile(file);
      setImportStatus(
        `Imported ${skills.length} skill${skills.length === 1 ? '' : 's'} from ${file.name}.`
      );
    } catch (error) {
      setImportError(error instanceof Error ? error.message : 'Failed to import skills ZIP.');
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
      const skill = await parseSkillMarkdownFile(file);
      setImportStatus(`Imported skill "${skill.metadata.name}" from ${file.name}.`);
    } catch (error) {
      setImportError(error instanceof Error ? error.message : 'Failed to import skill file.');
    } finally {
      setIsImporting(false);
      event.target.value = '';
    }
  }, []);

  return (
    <div className="space-y-3 p-3">
      {/* Import toolbar */}
      <div className="flex items-center justify-between gap-2">
        <h4 className="text-xs font-medium text-muted-foreground">Custom Skills</h4>
        <div className="flex items-center gap-1">
          <input
            ref={zipInputRef}
            type="file"
            accept=".zip,application/zip"
            className="hidden"
            aria-label="Import skills ZIP file"
            onChange={event => void handleImportZip(event)}
          />
          <input
            ref={mdInputRef}
            type="file"
            accept=".md,text/markdown"
            className="hidden"
            aria-label="Import skill Markdown file"
            onChange={event => void handleImportMd(event)}
          />
          <Button
            variant="secondary"
            size="sm"
            onClick={() => zipInputRef.current?.click()}
            disabled={isImporting}
            aria-busy={isImporting}
            title="Import skills from ZIP"
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
            title="Import a single skill .md file"
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

      {/* Bundled skills */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">Bundled (read-only)</p>
        {bundledSkills.length === 0 ? (
          <p className="text-xs text-muted-foreground">No bundled skills.</p>
        ) : (
          bundledSkills.map(skill => (
            <div
              key={`bundled-skill-${skill.metadata.name}`}
              className="flex items-center justify-between rounded-md border border-border px-2 py-1.5"
            >
              <div className="min-w-0">
                <p className="truncate text-sm font-medium">{skill.metadata.name}</p>
                <p className="truncate text-xs text-muted-foreground">
                  {skill.metadata.description}
                </p>
              </div>
              <Button
                variant="ghost"
                size="icon"
                className="size-7 shrink-0"
                onClick={() => downloadSkill(skill)}
                aria-label={`Download ${skill.metadata.name} as template`}
                title="Download as template"
              >
                <Codicon name="cloud-download" className="text-sm" />
              </Button>
            </div>
          ))
        )}
      </div>

      {/* Imported skills — managed via Plugin Hub */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">Imported</p>
        <p className="text-xs text-muted-foreground">No imported skills.</p>
      </div>
    </div>
  );
};
