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
      className="relative inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
      style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
      aria-label="MCP Tools"
      title="MCP Servers"
    >
      <Codicon name="extensions" className="text-[12px]" />
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
