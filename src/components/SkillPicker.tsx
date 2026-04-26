import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores/settingsStore';

export const SkillPicker: React.FC = () => {
  const [open, setOpen] = useState(false);
  const toggleSkill = useSettingsStore(s => s.toggleSkill);
  const disabledSkillNames = useSettingsStore(s => s.disabledSkillNames);

  const allSkills: { name: string; description: string }[] = [];
  const enabledCount = allSkills.filter(s => !disabledSkillNames.includes(s.name)).length;

  const renderSkillRow = (skill: { name: string; description: string }) => {
    const enabled = !disabledSkillNames.includes(skill.name);
    return (
      <button
        key={skill.name}
        onClick={() => toggleSkill(skill.name)}
        className={cn(
          'flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
          !enabled && 'opacity-50'
        )}
        aria-pressed={enabled}
        title={enabled ? `Disable ${skill.name}` : `Enable ${skill.name}`}
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
            <Codicon name="check" className="text-xs text-[var(--vscode-textLink-foreground)]" />
          )}
        </div>
        <div className="min-w-0 flex-1">
          <div className="font-medium text-foreground">{skill.name}</div>
          <div className="text-xs text-muted-foreground line-clamp-2">
            {skill.description.split('.')[0]}
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
            className="relative inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
            style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
            aria-label="Agent skills"
            title="Agent skills"
          >
            <Codicon name="lightbulb-sparkle" className="text-[14px]" />
            {enabledCount < allSkills.length && allSkills.length > 0 && (
              <span
                className="absolute -right-0.5 -top-0.5 flex h-3.5 min-w-3.5 items-center justify-center rounded-full bg-[var(--vscode-badge-background)] px-0.5 text-[8px] font-bold leading-none text-[var(--vscode-badge-foreground)]"
                aria-label={`${enabledCount} of ${allSkills.length} skills enabled`}
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
            {allSkills.length === 0 ? (
              <div className="px-2 py-2 text-xs text-muted-foreground space-y-1">
                <div>No skills available in the task pane.</div>
                <div>
                  Manage Copilot CLI plugins from your terminal with{' '}
                  <code className="font-mono">copilot plugin</code>.
                </div>
              </div>
            ) : (
              <>
                <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">Skills</div>
                {allSkills.map(renderSkillRow)}
              </>
            )}
          </Popover.Content>
        </Popover.Portal>
      </Popover.Root>
    </>
  );
};
