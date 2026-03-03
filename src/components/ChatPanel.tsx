import React from 'react';
import { Codicon } from '@/components/Codicon';
import { Thread } from '@/components/assistant-ui/thread';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';

interface ChatPanelProps {
  onOpenPanel?: (panel: string) => void;
}

/** VS Code-style icon-only tools button. */
const McpPill: React.FC<{ onOpenPanel?: (panel: string) => void }> = ({ onOpenPanel }) => {
  return (
    <button
      onClick={() => onOpenPanel?.('mcp')}
      className="relative inline-flex h-7 w-7 items-center justify-center rounded text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
      aria-label="MCP Tools"
      title="MCP Servers"
    >
      <Codicon name="extensions" className="text-base" />
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
            <ModelPicker />
          </>
        }
        rightToolbar={<McpPill onOpenPanel={onOpenPanel} />}
      />
    </div>
  );
};
