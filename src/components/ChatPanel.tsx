import React from 'react';
import { MessageList } from '@/components/chat/MessageList';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import { McpPicker } from './McpPicker';
import { Codicon } from './Codicon';
import type { McpOAuthPromptRequest } from './McpOAuthPrompt';
import type { ChatMessage } from '@/types';
import type { SessionMode } from '@/lib/websocket-client';

interface ChatPanelProps {
  messages: ChatMessage[];
  isRunning: boolean;
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  onSwitchModel?: (modelId: string) => Promise<void>;
  onSwitchAgent?: (agentName: string | null) => Promise<void>;
  onInitiateMcpOAuth?: (serverName: string, loginHint?: string) => Promise<string | undefined>;
  onOpenMcpOAuthPrompt?: (request: McpOAuthPromptRequest) => void;
  sessionMode: SessionMode;
  onSwitchSessionMode: (mode: SessionMode) => Promise<void>;
  onOpenPlan: () => void;
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
  onSwitchAgent,
  onInitiateMcpOAuth,
  onOpenMcpOAuthPrompt,
  sessionMode,
  onSwitchSessionMode,
  onOpenPlan,
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
            <AgentPicker onSwitchAgent={onSwitchAgent} />
            <ModelPicker hasActiveSession={messages.length > 0} onSwitchModel={onSwitchModel} />
            <button
              type="button"
              onClick={() =>
                void onSwitchSessionMode(sessionMode === 'plan' ? 'interactive' : 'plan')
              }
              className="aui-plan-mode-button inline-flex h-[22px] items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1.5 text-[11px] transition-colors hover:bg-[var(--vscode-toolbar-hoverBackground)]"
              style={{
                color:
                  sessionMode === 'plan'
                    ? 'var(--vscode-badge-foreground)'
                    : 'var(--vscode-icon-foreground)',
                backgroundColor:
                  sessionMode === 'plan' ? 'var(--vscode-badge-background)' : 'transparent',
              }}
              title={sessionMode === 'plan' ? 'Exit Plan mode' : 'Enter Plan mode'}
              aria-label={sessionMode === 'plan' ? 'Exit Plan mode' : 'Enter Plan mode'}
            >
              <Codicon name="checklist" className="text-xs" />
              Plan
            </button>
            <button
              type="button"
              onClick={onOpenPlan}
              className="inline-flex h-[22px] items-center justify-center rounded-[var(--vscode-cornerRadius-small)] px-1 transition-colors hover:bg-[var(--vscode-toolbar-hoverBackground)]"
              style={{ color: 'var(--vscode-icon-foreground)' }}
              title="Open plan"
              aria-label="Open plan"
            >
              <Codicon name="note" className="text-xs" />
            </button>
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
