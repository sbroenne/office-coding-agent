import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { BrainCircuit, Check, Square, ChevronDown } from 'lucide-react';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import { getSkills } from '@/services/skills';

export const SkillPicker: React.FC = () => {
  const [open, setOpen] = useState(false);
  const activeSkillNames = useSettingsStore(s => s.activeSkillNames);
  const toggleSkill = useSettingsStore(s => s.toggleSkill);

  const allSkills = getSkills();

  if (allSkills.length === 0) return null;

  // null = all on, explicit array = only those
  const allNames = allSkills.map(s => s.metadata.name);
  const effectiveActive: string[] = activeSkillNames ?? allNames;

  const activeCount = effectiveActive.filter(n =>
    allSkills.some(s => s.metadata.name === n)
  ).length;

  return (
    <Popover.Root open={open} onOpenChange={setOpen}>
      <Popover.Trigger asChild>
        <button
          className="relative inline-flex items-center gap-1 rounded-md px-2 py-1 text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
          aria-label="Agent skills"
          title="Agent skills"
        >
          <BrainCircuit className="size-4" />
          {activeCount > 0 && (
            <span className="inline-flex h-4 min-w-4 items-center justify-center rounded-full bg-primary px-1 text-[10px] font-medium text-primary-foreground">
              {activeCount}
            </span>
          )}
          <ChevronDown className="size-3 opacity-60" />
        </button>
      </Popover.Trigger>

      <Popover.Portal>
        <Popover.Content
          className="z-50 w-64 max-h-80 overflow-y-auto rounded-lg border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
          sideOffset={4}
          align="start"
        >
          <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">Skills</div>
          {allSkills.map(skill => {
            const isActive = effectiveActive.includes(skill.metadata.name);
            return (
              <button
                key={skill.metadata.name}
                onClick={() => toggleSkill(skill.metadata.name)}
                className="flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent"
              >
                <div className="mt-0.5 flex size-4 shrink-0 items-center justify-center rounded border border-border">
                  {isActive ? (
                    <Check className="size-3 text-primary" />
                  ) : (
                    <Square className="size-3 opacity-0" />
                  )}
                </div>
                <div className="min-w-0 flex-1">
                  <div
                    className={cn(
                      'font-medium',
                      isActive ? 'text-foreground' : 'text-muted-foreground'
                    )}
                  >
                    {skill.metadata.name}
                  </div>
                  <div className="text-xs text-muted-foreground line-clamp-2">
                    {skill.metadata.description.split('.')[0]}
                  </div>
                </div>
              </button>
            );
          })}
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
