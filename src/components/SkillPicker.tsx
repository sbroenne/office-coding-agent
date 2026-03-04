import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import { getBundledSkills, getImportedSkills, getSkills } from '@/services/skills';
import { detectOfficeHost } from '@/services/office/host';

interface SkillPickerProps {
  onOpenPanel?: (panel: string) => void;
}

export const SkillPicker: React.FC<SkillPickerProps> = ({ onOpenPanel }) => {
  const [open, setOpen] = useState(false);
  const activeSkillNames = useSettingsStore(s => s.activeSkillNames);
  const toggleSkill = useSettingsStore(s => s.toggleSkill);

  const host = detectOfficeHost();
  const allSkills = getSkills().filter(
    s => s.metadata.hosts.length === 0 || (host !== 'unknown' && s.metadata.hosts.includes(host))
  );
  const bundledSkills = getBundledSkills().filter(
    s => s.metadata.hosts.length === 0 || (host !== 'unknown' && s.metadata.hosts.includes(host))
  );
  const importedSkills = getImportedSkills().filter(
    s => s.metadata.hosts.length === 0 || (host !== 'unknown' && s.metadata.hosts.includes(host))
  );

  // null = all on, explicit array = only those
  const allNames = allSkills.map(s => s.metadata.name);
  const effectiveActive: string[] = activeSkillNames ?? allNames;

  const activeCount = effectiveActive.filter(n =>
    allSkills.some(s => s.metadata.name === n)
  ).length;

  const renderSkillOption = (skillName: string, skillDescription: string) => {
    const isActive = effectiveActive.includes(skillName);

    return (
      <button
        key={skillName}
        onClick={() => toggleSkill(skillName)}
        className="flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent"
      >
        <div className="mt-0.5 flex size-4 shrink-0 items-center justify-center rounded border border-border">
          {isActive ? (
            <Codicon name="check" className="text-xs text-primary" />
          ) : (
            <Codicon name="primitive-square" className="text-xs opacity-0" />
          )}
        </div>
        <div className="min-w-0 flex-1">
          <div
            className={cn('font-medium', isActive ? 'text-foreground' : 'text-muted-foreground')}
          >
            {skillName}
          </div>
          <div className="text-xs text-muted-foreground line-clamp-2">
            {skillDescription.split('.')[0]}
          </div>
        </div>
      </button>
    );
  };

  return (
    <>
      <Popover.Root open={open} onOpenChange={setOpen}>
        <Popover.Trigger asChild>
          <button
            className="relative inline-flex h-8 w-8 items-center justify-center rounded text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
            aria-label="Agent skills"
            title="Agent skills"
          >
            <Codicon name="lightbulb-sparkle" className="text-base" />
            {activeCount > 0 && activeSkillNames !== null && (
              <span className="absolute -top-0.5 -right-0.5 flex h-3.5 min-w-[14px] items-center justify-center rounded-full bg-primary px-0.5 text-[9px] font-medium tabular-nums text-primary-foreground">
                {activeCount}
              </span>
            )}
          </button>
        </Popover.Trigger>

        <Popover.Portal>
          <Popover.Content
            className="z-50 w-64 max-h-80 overflow-y-auto rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
            sideOffset={4}
            align="start"
          >
            <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">Skills</div>
            {bundledSkills.length > 0 && (
              <>
                <div className="flex items-center justify-between px-2 py-1 text-[10px] uppercase tracking-wide text-muted-foreground">
                  <span>Bundled</span>
                  <span>Read-only category</span>
                </div>
                {bundledSkills.map(skill =>
                  renderSkillOption(skill.metadata.name, skill.metadata.description)
                )}
              </>
            )}

            {importedSkills.length > 0 && (
              <>
                <div className="mt-1 flex items-center justify-between px-2 py-1 text-[10px] uppercase tracking-wide text-muted-foreground">
                  <span>Imported</span>
                  <span>ZIP</span>
                </div>
                {importedSkills.map(skill =>
                  renderSkillOption(skill.metadata.name, skill.metadata.description)
                )}
              </>
            )}

            {allSkills.length === 0 && (
              <div className="px-2 py-2 text-xs text-muted-foreground">
                No skills available yet.
              </div>
            )}

            <div className="mt-1 border-t border-border pt-1">
              <button
                onClick={() => {
                  setOpen(false);
                  onOpenPanel?.('skills');
                }}
                className="flex w-full items-center justify-between rounded-md px-2 py-1.5 text-left text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
              >
                <span>Manage skills…</span>
                <span>{importedSkills.length} imported</span>
              </button>
            </div>
          </Popover.Content>
        </Popover.Portal>
      </Popover.Root>
    </>
  );
};
