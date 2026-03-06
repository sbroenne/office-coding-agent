import React from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import { getBundledAgents } from '@/services/agents';
import { downloadAgent } from '@/services/extensions/zipExportService';

export const AgentManagerPanel: React.FC = () => {
  const bundledAgents = getBundledAgents();

  return (
    <div className="space-y-3 p-3">
      {/* Bundled agents */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">Bundled (read-only)</p>
        {bundledAgents.length === 0 ? (
          <p className="text-xs text-muted-foreground">No bundled agents.</p>
        ) : (
          bundledAgents.map(agent => (
            <div
              key={`bundled-agent-${agent.metadata.name}`}
              className="flex items-center justify-between rounded-md border border-border px-2 py-1.5"
            >
              <div className="min-w-0">
                <p className="truncate text-sm font-medium">{agent.metadata.name}</p>
                <p className="truncate text-xs text-muted-foreground">
                  {agent.metadata.description}
                </p>
              </div>
              <Button
                variant="ghost"
                size="icon"
                className="size-7 shrink-0"
                onClick={() => downloadAgent(agent)}
                aria-label={`Download ${agent.metadata.name} as template`}
                title="Download as template"
              >
                <Codicon name="cloud-download" className="text-sm" />
              </Button>
            </div>
          ))
        )}
      </div>

      {/* Plugin agents */}
      <div className="space-y-1">
        <p className="text-[11px] font-medium text-muted-foreground">From plugins</p>
        <p className="text-xs text-muted-foreground">
          Custom agents are installed via the Copilot CLI Plugin Hub.
        </p>
      </div>
    </div>
  );
};
