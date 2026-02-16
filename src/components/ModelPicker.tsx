import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Check, ChevronDown, Settings } from 'lucide-react';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import type { ModelInfo, ModelProvider } from '@/types';

/** Provider display order */
const PROVIDER_ORDER: ModelProvider[] = [
  'Anthropic',
  'OpenAI',
  'DeepSeek',
  'Meta',
  'Mistral',
  'xAI',
  'Microsoft',
  'Other',
];

interface ModelPickerProps {
  onOpenSettings?: () => void;
}

export const ModelPicker: React.FC<ModelPickerProps> = ({ onOpenSettings }) => {
  const [open, setOpen] = useState(false);

  const {
    activeModelId,
    setActiveModel,
    getActiveEndpoint,
    getActiveModel,
    getModelsForActiveEndpoint,
  } = useSettingsStore();

  const activeEndpoint = getActiveEndpoint();
  const activeModel = getActiveModel();
  const models = getModelsForActiveEndpoint();

  const groupedModels = groupByProvider(models);
  const displayLabel = activeModel?.name ?? 'Select model';

  if (!activeEndpoint) {
    return (
      <button
        className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs text-muted-foreground opacity-50 cursor-not-allowed"
        disabled
      >
        No endpoint configured
      </button>
    );
  }

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
          className="z-50 w-64 max-h-80 overflow-y-auto rounded-lg border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
          sideOffset={4}
          align="start"
        >
          {models.length === 0 && (
            <button
              onClick={() => {
                onOpenSettings?.();
                setOpen(false);
              }}
              className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-sm text-muted-foreground transition-colors hover:bg-accent"
            >
              <Settings className="size-3.5" />
              Configure models in Settings
            </button>
          )}

          {PROVIDER_ORDER.filter(p => groupedModels.get(p)?.length).map((provider, idx, arr) => {
            const providerModels = groupedModels.get(provider) ?? [];
            return (
              <div key={provider}>
                <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">
                  {provider}
                </div>
                {providerModels.map(model => {
                  const isActive = model.id === activeModelId;
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

function groupByProvider(models: ModelInfo[]): Map<ModelProvider, ModelInfo[]> {
  const groups = new Map<ModelProvider, ModelInfo[]>();

  for (const model of models) {
    const provider = model.provider;
    const group = groups.get(provider) ?? [];
    group.push(model);
    groups.set(provider, group);
  }

  return groups;
}
