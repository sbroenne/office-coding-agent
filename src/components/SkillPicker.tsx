import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { getBundledSkills, getSkills } from '@/services/skills';
import { detectOfficeHost } from '@/services/office/host';

interface SkillPickerProps {
  onOpenPanel?: (panel: string) => void;
}

export const SkillPicker: React.FC<SkillPickerProps> = ({ onOpenPanel }) => {
  const [open, setOpen] = useState(false);

  const host = detectOfficeHost();
  const allSkills = getSkills().filter(
    s => s.metadata.hosts.length === 0 || (host !== 'unknown' && s.metadata.hosts.includes(host))
  );
  const bundledSkills = getBundledSkills().filter(
    s => s.metadata.hosts.length === 0 || (host !== 'unknown' && s.metadata.hosts.includes(host))
  );

  const renderSkillOption = (skillName: string, skillDescription: string) => {
    return (
      <div
        key={skillName}
        className="flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm"
      >
        <div className="mt-0.5 flex size-4 shrink-0 items-center justify-center rounded border border-border">
          <Codicon name="check" className="text-xs text-primary" />
        </div>
        <div className="min-w-0 flex-1">
          <div className={cn('font-medium text-foreground')}>{skillName}</div>
          <div className="text-xs text-muted-foreground line-clamp-2">
            {skillDescription.split('.')[0]}
          </div>
        </div>
      </div>
    );
  };

  return (
    <>
      <Popover.Root open={open} onOpenChange={setOpen}>
        <Popover.Trigger asChild>
          <button
            className="relative inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
            style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
            aria-label="Agent skills"
            title="Agent skills"
          >
            <Codicon name="lightbulb-sparkle" className="text-[14px]" />
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
                  <span>Read-only</span>
                </div>
                {bundledSkills.map(skill =>
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
                  onOpenPanel?.('plugins');
                }}
                className="flex w-full items-center justify-between rounded-md px-2 py-1.5 text-left text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
              >
                <span>Manage plugins…</span>
              </button>
            </div>
          </Popover.Content>
        </Popover.Portal>
      </Popover.Root>
    </>
  );
};
