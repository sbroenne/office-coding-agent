import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Check, ChevronDown } from 'lucide-react';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import { COPILOT_MODELS } from '@/types';
import type { CopilotModel } from '@/types';

const PROVIDER_ORDER: CopilotModel['provider'][] = ['Anthropic', 'OpenAI', 'Google', 'Other'];

export const ModelPicker: React.FC = () => {
  const [open, setOpen] = useState(false);
  const { activeModel, setActiveModel } = useSettingsStore();

  const currentModel = COPILOT_MODELS.find(m => m.id === activeModel);
  const displayLabel = currentModel?.name ?? 'Select model';

  const groupedModels = COPILOT_MODELS.reduce((groups, model) => {
    const group = groups.get(model.provider) ?? [];
    group.push(model);
    groups.set(model.provider, group);
    return groups;
  }, new Map<CopilotModel['provider'], CopilotModel[]>());

  return (
    <Popover.Root open={open} onOpenChange={setOpen}>
      <Popover.Trigger asChild>
        <button
          className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
          aria-label="Select model"
          title="Select model"
        >
          <span className="max-w-[120px] truncate">{displayLabel}</span>
          <ChevronDown className="size-3 opacity-60" />
        </button>
      </Popover.Trigger>

      <Popover.Portal>
        <Popover.Content
          className="z-50 w-64 max-h-80 overflow-y-auto rounded-lg border border-border bg-popover p-1 text-popover-foreground shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
          sideOffset={4}
          align="start"
        >
          {PROVIDER_ORDER.filter(p => groupedModels.get(p)?.length).map((provider, idx, arr) => {
            const providerModels = groupedModels.get(provider) ?? [];
            return (
              <div key={provider}>
                <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">
                  {provider}
                </div>
                {providerModels.map(model => {
                  const isActive = model.id === activeModel;
                  return (
                    <button
                      key={model.id}
                      onClick={() => {
                        setActiveModel(model.id);
                        setOpen(false);
                      }}
                      className={cn(
                        'flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
                        isActive && 'bg-accent/50'
                      )}
                    >
                      <Check
                        className={cn('size-3.5 shrink-0', isActive ? 'opacity-100' : 'opacity-0')}
                      />
                      <span className="truncate text-foreground">{model.name}</span>
                    </button>
                  );
                })}
                {idx < arr.length - 1 && <div className="my-1 h-px bg-border" />}
              </div>
            );
          })}
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
