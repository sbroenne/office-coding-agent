import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';

interface AgentPickerProps {
  onSwitchAgent?: (agentName: string | null) => Promise<void>;
}

export const AgentPicker: React.FC<AgentPickerProps> = ({ onSwitchAgent }) => {
  const [open, setOpen] = useState(false);
  const [isSwitching, setIsSwitching] = useState(false);
  const [switchError, setSwitchError] = useState<string | null>(null);
  const { activeAgentName, availableAgents, setActiveAgent } = useSettingsStore();

  const agents = availableAgents ?? [];
  const activeAgent = activeAgentName ? agents.find(agent => agent.name === activeAgentName) : null;
  const displayName = activeAgent?.displayName ?? 'Default';

  const selectAgent = (agentName: string | null) => {
    void (async () => {
      setSwitchError(null);
      setIsSwitching(true);
      try {
        if (onSwitchAgent) {
          await onSwitchAgent(agentName);
        } else {
          setActiveAgent(agentName);
        }
        setOpen(false);
      } catch (err) {
        setSwitchError(err instanceof Error ? err.message : 'Failed to switch agent');
      } finally {
        setIsSwitching(false);
      }
    })();
  };

  const renderAgentOption = (
    agentName: string | null,
    agentLabel: string,
    agentDescription: string,
    isActive: boolean
  ) => (
    <button
      key={agentName ?? 'default'}
      type="button"
      onClick={() => selectAgent(agentName)}
      disabled={isSwitching}
      className={cn(
        'flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
        isActive && 'bg-accent/50',
        isSwitching && 'opacity-50 cursor-not-allowed'
      )}
    >
      <Codicon
        name="check"
        className={cn('mt-0.5 text-[12px] shrink-0', isActive ? 'opacity-100' : 'opacity-0')}
      />
      <div className="min-w-0 flex-1">
        <div className="truncate font-medium text-foreground">{agentLabel}</div>
        <div className="line-clamp-2 text-xs text-muted-foreground">{agentDescription}</div>
      </div>
    </button>
  );

  return (
    <Popover.Root open={open} onOpenChange={setOpen}>
      <Popover.Trigger asChild>
        <button
          className="relative inline-flex h-7 items-center gap-1 rounded px-1.5 text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
          aria-label="Select agent"
          title={`Agent: ${displayName}`}
        >
          <Codicon name="robot" className="text-base" />
          <span className="max-w-[110px] truncate text-xs">{displayName}</span>
          <Codicon name="chevron-down" className="text-[12px] shrink-0 opacity-60" />
        </button>
      </Popover.Trigger>

      <Popover.Portal>
        <Popover.Content
          className="z-50 w-64 max-h-80 overflow-y-auto rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
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
          <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">Agents</div>
          {renderAgentOption(
            null,
            'Default',
            'Use the default Copilot CLI agent for this Office host.',
            activeAgentName === null
          )}
          {availableAgents === null ? (
            <div className="px-3 py-3 text-center text-xs text-muted-foreground">
              Connecting to Copilot…
            </div>
          ) : agents.length === 0 ? (
            <div className="px-3 py-3 text-center text-xs text-muted-foreground">
              No CLI agents found.
            </div>
          ) : (
            agents.map(agent =>
              renderAgentOption(
                agent.name,
                agent.displayName,
                agent.description.split('.')[0] || agent.description,
                agent.name === activeAgentName
              )
            )
          )}
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
