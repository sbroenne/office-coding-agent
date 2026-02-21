import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Check, ChevronDown } from 'lucide-react';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import type { CopilotModel, ModelProvider } from '@/types';

const PROVIDER_ORDER: ModelProvider[] = ['Anthropic', 'OpenAI', 'Google', 'Other'];

/** Convert a raw model ID like 'claude-sonnet-4' to 'Claude Sonnet 4' */
function formatModelId(id: string): string {
  return id
    .split('-')
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(' ');
}

export const ModelPicker: React.FC = () => {
  const [open, setOpen] = useState(false);
  const { activeModel, setActiveModel, availableModels } = useSettingsStore();

  const models = availableModels ?? [];
  const currentModel = models.find(m => m.id === activeModel);
  const displayLabel = currentModel?.name ?? formatModelId(activeModel);

  const groupedModels = models.reduce((groups, model) => {
    const group = groups.get(model.provider) ?? [];
    group.push(model);
    groups.set(model.provider, group);
    return groups;
  }, new Map<ModelProvider, CopilotModel[]>());

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
          {models.length === 0 ? (
            <div className="px-3 py-4 text-center text-xs text-muted-foreground">
              Connecting to Copilot…
            </div>
          ) : (
            PROVIDER_ORDER.filter(p => groupedModels.get(p)?.length).map((provider, idx, arr) => {
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
                          className={cn(
                            'size-3.5 shrink-0',
                            isActive ? 'opacity-100' : 'opacity-0'
                          )}
                        />
                        <span className="truncate text-foreground">{model.name}</span>
                      </button>
                    );
                  })}
                  {idx < arr.length - 1 && <div className="my-1 h-px bg-border" />}
                </div>
              );
            })
          )}
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
