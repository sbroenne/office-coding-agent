import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { Bot, Check, ChevronDown } from 'lucide-react';
import { cn } from '@/lib/utils';
import { useSettingsStore } from '@/stores';
import {
  getAgents,
  getBundledAgents,
  getImportedAgents,
  resolveActiveAgent,
} from '@/services/agents';
import { detectOfficeHost } from '@/services/office/host';
import { AgentManagerDialog } from './AgentManagerDialog';

export const AgentPicker: React.FC = () => {
  const [open, setOpen] = useState(false);
  const [managerOpen, setManagerOpen] = useState(false);
  const activeAgentId = useSettingsStore(s => s.activeAgentId);
  const setActiveAgent = useSettingsStore(s => s.setActiveAgent);

  const host = detectOfficeHost();
  const targetHost = host === 'excel' || host === 'powerpoint' ? host : undefined;
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

  const renderAgentOption = (agentName: string, agentDescription: string) => {
    const isActive = agentName === activeAgentId;

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
        <Check className={cn('mt-0.5 size-3.5 shrink-0', isActive ? 'opacity-100' : 'opacity-0')} />
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
            className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
            aria-label="Select agent"
            title="Select agent"
          >
            <Bot className="size-3.5" />
            <span className="max-w-[100px] truncate">{displayName}</span>
            <ChevronDown className="size-3 opacity-60" />
          </button>
        </Popover.Trigger>

        <Popover.Portal>
          <Popover.Content
            className="z-50 w-56 rounded-lg border border-border bg-popover p-1 shadow-md outline-none animate-in fade-in-0 zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=top]:slide-in-from-bottom-2"
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
                  setManagerOpen(true);
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

      <AgentManagerDialog open={managerOpen} onOpenChange={setManagerOpen} />
    </>
  );
};
