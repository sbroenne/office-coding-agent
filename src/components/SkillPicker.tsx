import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { getBundledSkills } from '@/services/skills';
import { detectOfficeHost } from '@/services/office/host';
import { useSettingsStore } from '@/stores/settingsStore';

interface SkillPickerProps {
  onOpenPanel?: (panel: string) => void;
}

export const SkillPicker: React.FC<SkillPickerProps> = ({ onOpenPanel }) => {
  const [open, setOpen] = useState(false);
  const toggleSkill = useSettingsStore(s => s.toggleSkill);
  const disabledSkillNames = useSettingsStore(s => s.disabledSkillNames);

  const host = detectOfficeHost();
  const bundledSkills = getBundledSkills().filter(
    s => s.metadata.hosts.length === 0 || (host !== 'unknown' && s.metadata.hosts.includes(host))
  );

  const enabledCount = bundledSkills.filter(
    s => !disabledSkillNames.includes(s.metadata.name)
  ).length;

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
            {enabledCount < bundledSkills.length && bundledSkills.length > 0 && (
              <span
                className="absolute -right-0.5 -top-0.5 flex h-3.5 min-w-3.5 items-center justify-center rounded-full bg-[var(--vscode-badge-background)] px-0.5 text-[8px] font-bold leading-none text-[var(--vscode-badge-foreground)]"
                aria-label={`${enabledCount} of ${bundledSkills.length} skills enabled`}
              >
                {enabledCount}
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

            {bundledSkills.length === 0 ? (
              <div className="px-2 py-2 text-xs text-muted-foreground">
                No skills available yet.
              </div>
            ) : (
              bundledSkills.map(skill => {
                const enabled = !disabledSkillNames.includes(skill.metadata.name);
                return (
                  <button
                    key={skill.metadata.name}
                    onClick={() => toggleSkill(skill.metadata.name)}
                    className={cn(
                      'flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
                      !enabled && 'opacity-50'
                    )}
                    aria-pressed={enabled}
                    title={
                      enabled ? `Disable ${skill.metadata.name}` : `Enable ${skill.metadata.name}`
                    }
                  >
                    <div
                      className={cn(
                        'mt-0.5 flex size-4 shrink-0 items-center justify-center rounded border',
                        enabled
                          ? 'border-[var(--vscode-textLink-foreground)] bg-[var(--vscode-textLink-foreground)]/10'
                          : 'border-border'
                      )}
                    >
                      {enabled && (
                        <Codicon
                          name="check"
                          className="text-xs text-[var(--vscode-textLink-foreground)]"
                        />
                      )}
                    </div>
                    <div className="min-w-0 flex-1">
                      <div className="font-medium text-foreground">{skill.metadata.name}</div>
                      <div className="text-xs text-muted-foreground line-clamp-2">
                        {skill.metadata.description.split('.')[0]}
                      </div>
                    </div>
                  </button>
                );
              })
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
