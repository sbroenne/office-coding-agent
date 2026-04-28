import React from 'react';
import { MessageList } from '@/components/chat/MessageList';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import { McpPicker } from './McpPicker';
import type { McpOAuthPromptRequest } from './McpOAuthPrompt';
import type { ChatMessage } from '@/types';

interface ChatPanelProps {
  messages: ChatMessage[];
  isRunning: boolean;
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  onSwitchModel?: (modelId: string) => Promise<void>;
  onInitiateMcpOAuth?: (serverName: string, loginHint?: string) => Promise<string | undefined>;
  onOpenMcpOAuthPrompt?: (request: McpOAuthPromptRequest) => void;
  onEnqueue?: (text: string) => void;
  queuedPrompts?: string[];
  onDequeue?: (index: number) => void;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({
  messages,
  isRunning,
  onSend,
  onCancel,
  onSwitchModel,
  onInitiateMcpOAuth,
  onOpenMcpOAuthPrompt,
  onEnqueue,
  queuedPrompts,
  onDequeue,
}) => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      <MessageList
        messages={messages}
        isRunning={isRunning}
        onSend={onSend}
        onCancel={onCancel}
        onEnqueue={onEnqueue}
        queuedPrompts={queuedPrompts}
        onDequeue={onDequeue}
        onFeedback={() => {
          /* TODO */
        }}
        onRegenerate={() => {
          /* TODO */
        }}
        leftToolbar={
          <>
            <AgentPicker />
            <ModelPicker hasActiveSession={messages.length > 0} onSwitchModel={onSwitchModel} />
          </>
        }
        rightToolbar={
          <McpPicker
            onInitiateOAuth={onInitiateMcpOAuth}
            onOpenOAuthPrompt={onOpenMcpOAuthPrompt}
          />
        }
      />
    </div>
  );
};
