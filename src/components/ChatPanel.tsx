import React from 'react';
import { Codicon } from '@/components/Codicon';
import { MessageList } from '@/components/chat/MessageList';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import type { ChatMessage } from '@/types';

interface ChatPanelProps {
  messages: ChatMessage[];
  isRunning: boolean;
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
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

export const ChatPanel: React.FC<ChatPanelProps> = ({
  messages,
  isRunning,
  onSend,
  onCancel,
  onOpenPanel,
}) => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      <MessageList
        messages={messages}
        isRunning={isRunning}
        onSend={onSend}
        onCancel={onCancel}
        onFeedback={() => {/* TODO */}}
        onRegenerate={() => {/* TODO */}}
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
