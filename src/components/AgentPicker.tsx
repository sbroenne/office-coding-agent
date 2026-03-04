import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Codicon } from '@/components/Codicon';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import {
  getAgents,
  getBundledAgents,
  getImportedAgents,
  resolveActiveAgent,
} from '@/services/agents';
import { detectOfficeHost } from '@/services/office/host';

interface AgentPickerProps {
  onOpenPanel?: (panel: string) => void;
}

export const AgentPicker: React.FC<AgentPickerProps> = ({ onOpenPanel }) => {
  const [open, setOpen] = useState(false);
  const activeAgentId = useSettingsStore(s => s.activeAgentId);
  const setActiveAgent = useSettingsStore(s => s.setActiveAgent);

  const host = detectOfficeHost();
  const targetHost =
    host === 'excel' || host === 'powerpoint' || host === 'word' || host === 'outlook'
      ? host
      : undefined;
  const allAgents = getAgents(host);
  const bundledAgents = targetHost
    ? getBundledAgents().filter(agent => agent.metadata.hosts.includes(targetHost))
    : [];
  const importedAgents = targetHost
    ? getImportedAgents().filter(agent => agent.metadata.hosts.includes(targetHost))
    : [];

  if (allAgents.length === 0) return null;

  const activeAgent = resolveActiveAgent(activeAgentId, host);
  const displayName = activeAgent?.metadata.name ?? allAgents[0].metadata.name;

  const resolvedName = activeAgent?.metadata.name ?? null;

  const renderAgentOption = (agentName: string, agentDescription: string) => {
    const isActive = agentName === resolvedName;

    return (
      <button
        key={agentName}
        onClick={() => {
          setActiveAgent(agentName);
          setOpen(false);
        }}
        className={cn(
          'flex w-full items-start gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent',
          isActive && 'bg-accent/50'
        )}
      >
        <Codicon
          name="check"
          className={cn('mt-0.5 text-sm shrink-0', isActive ? 'opacity-100' : 'opacity-0')}
        />
        <div className="min-w-0 flex-1">
          <div className="font-medium text-foreground">{agentName}</div>
          <div className="text-xs text-muted-foreground line-clamp-2">
            {agentDescription.split('.')[0]}
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
            className="relative inline-flex items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1.5 transition-colors hover:bg-accent"
            style={{ height: 22, color: 'var(--vscode-icon-foreground)', fontSize: 12 }}
            aria-label="Select agent"
            title={`Agent: ${displayName}`}
          >
            <Codicon name="robot" className="text-[12px]" />
            <span className="truncate max-w-[80px]">{displayName}</span>
          </button>
        </Popover.Trigger>

        <Popover.Portal>
          <Popover.Content
            className="z-50 w-56 rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
            sideOffset={4}
            align="start"
          >
            <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">Agent</div>
            {bundledAgents.length > 0 && (
              <>
                <div className="flex items-center justify-between px-2 py-1 text-[10px] uppercase tracking-wide text-muted-foreground">
                  <span>Bundled</span>
                  <span>Read-only</span>
                </div>
                {bundledAgents.map(agent =>
                  renderAgentOption(agent.metadata.name, agent.metadata.description)
                )}
              </>
            )}

            {importedAgents.length > 0 && (
              <>
                <div className="mt-1 flex items-center justify-between px-2 py-1 text-[10px] uppercase tracking-wide text-muted-foreground">
                  <span>Imported</span>
                  <span>ZIP</span>
                </div>
                {importedAgents.map(agent =>
                  renderAgentOption(agent.metadata.name, agent.metadata.description)
                )}
              </>
            )}

            <div className="mt-1 border-t border-border pt-1">
              <button
                onClick={() => {
                  setOpen(false);
                  onOpenPanel?.('agents');
                }}
                className="flex w-full items-center justify-between rounded-md px-2 py-1.5 text-left text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
              >
                <span>Manage agents…</span>
                <span>{importedAgents.length} imported</span>
              </button>
            </div>
          </Popover.Content>
        </Popover.Portal>
      </Popover.Root>
    </>
  );
};
