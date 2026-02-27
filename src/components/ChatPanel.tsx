import React from 'react';
import { Server } from 'lucide-react';
import { Thread } from '@/components/assistant-ui/thread';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import { SkillPicker } from './SkillPicker';
import { useSettingsStore } from '@/stores';
import { BUNDLED_MCP_SERVERS } from '@/types';

interface ChatPanelProps {
  onOpenPanel?: (panel: string) => void;
}

/** VS Code-style "Tools N" pill — shows active MCP server count, opens MCP panel. */
const McpPill: React.FC<{ onOpenPanel?: (panel: string) => void }> = ({ onOpenPanel }) => {
  const activeMcpServerNames = useSettingsStore(s => s.activeMcpServerNames);
  const importedMcpServers = useSettingsStore(s => s.importedMcpServers);

  const allServers = [...BUNDLED_MCP_SERVERS, ...importedMcpServers];
  const activeCount = allServers.filter(
    s => activeMcpServerNames === null || activeMcpServerNames.includes(s.name)
  ).length;

  return (
    <button
      onClick={() => onOpenPanel?.('mcp')}
      className="inline-flex items-center gap-1.5 rounded-full border border-border/60 py-0.5 pl-1.5 pr-2 text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
      aria-label="MCP Tools"
      title="MCP Servers"
    >
      <Server className="size-3 shrink-0" />
      <span>Tools</span>
      {activeCount > 0 && (
        <span className="min-w-[14px] rounded-full bg-muted px-1 text-center text-[10px] tabular-nums text-muted-foreground">
          {activeCount}
        </span>
      )}
    </button>
  );
};

export const ChatPanel: React.FC<ChatPanelProps> = ({ onOpenPanel }) => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      <Thread
        leftToolbar={
          <>
            <AgentPicker onOpenPanel={onOpenPanel} />
            <SkillPicker onOpenPanel={onOpenPanel} />
            <McpPill onOpenPanel={onOpenPanel} />
          </>
        }
        rightToolbar={<ModelPicker />}
      />
    </div>
  );
};
