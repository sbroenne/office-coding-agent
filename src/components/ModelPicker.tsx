import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
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

interface ModelPickerProps {
  hasActiveSession?: boolean;
  onSwitchModel?: (modelId: string) => Promise<void>;
}

export const ModelPicker: React.FC<ModelPickerProps> = ({
  hasActiveSession = false,
  onSwitchModel,
}) => {
  const [open, setOpen] = useState(false);
  const [isSwitching, setIsSwitching] = useState(false);
  const [switchError, setSwitchError] = useState<string | null>(null);
  const { activeModel, setActiveModel, availableModels } = useSettingsStore();

  const models = availableModels ?? [];
  const currentModel = models.find(m => m.id === activeModel);
  const displayLabel = isSwitching
    ? '(switching…)'
    : (currentModel?.name ?? formatModelId(activeModel));

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
          className="inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1.5 text-[12px] transition-colors hover:bg-accent"
          style={{ height: 22, color: 'var(--vscode-icon-foreground)' }}
          aria-label="Select model"
          title="Select model"
        >
          <span className="max-w-[110px] truncate">{displayLabel}</span>
          <Codicon name="chevron-down" className="text-[12px] shrink-0 opacity-60" />
        </button>
      </Popover.Trigger>

      <Popover.Portal>
        <Popover.Content
          className="z-50 w-64 max-h-80 overflow-y-auto rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-popover p-1 text-popover-foreground shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
          sideOffset={4}
          align="start"
        >
          {switchError && (
            <div
              className="border-b border-border px-3 py-2 text-xs"
              style={{ color: 'var(--vscode-errorForeground)' }}
            >
              {switchError}
            </div>
          )}
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
                          void (async () => {
                            setSwitchError(null);
                            if (hasActiveSession && onSwitchModel) {
                              setIsSwitching(true);
                              try {
                                await onSwitchModel(model.id);
                                setOpen(false);
                              } catch (err) {
                                setSwitchError(
                                  err instanceof Error ? err.message : 'Failed to switch model'
                                );
                              } finally {
                                setIsSwitching(false);
                              }
                            } else {
                              setActiveModel(model.id);
                              setOpen(false);
                            }
                          })();
                        }}
                        disabled={isSwitching}
                        className={cn(
                          'flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
                          isActive && 'bg-accent/50',
                          isSwitching && 'opacity-50 cursor-not-allowed'
                        )}
                      >
                        <Codicon
                          name="check"
                          className={cn('text-sm shrink-0', isActive ? 'opacity-100' : 'opacity-0')}
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
