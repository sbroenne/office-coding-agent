import React, { useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { MessageList } from '@/components/chat/MessageList';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import { McpPicker } from './McpPicker';
import { Codicon } from './Codicon';
import type { McpOAuthPromptRequest } from './McpOAuthPrompt';
import type { ChatMessage } from '@/types';
import type { SessionMode } from '@/lib/websocket-client';
import { detectOfficeHost, type OfficeHostApp } from '@/services/office/host';

const CONVERSATION_MODES: { mode: SessionMode; label: string; description: string }[] = [
  {
    mode: 'interactive',
    label: 'Interactive',
    description: 'Work together, one request at a time.',
  },
  { mode: 'plan', label: 'Plan', description: 'Make a plan before changing your document.' },
  {
    mode: 'autopilot',
    label: 'Autopilot',
    description: 'Continue working until the task is complete. Permission approvals still apply.',
  },
];

interface ChatPanelProps {
  host?: OfficeHostApp;
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
  host = detectOfficeHost(),
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
  const [isSwitchingMode, setIsSwitchingMode] = useState(false);
  const [modeError, setModeError] = useState<string | null>(null);
  const [modePickerOpen, setModePickerOpen] = useState(false);
  const modeLabel = CONVERSATION_MODES.find(option => option.mode === sessionMode)?.label;
  const selectMode = async (mode: SessionMode) => {
    setModeError(null);
    setIsSwitchingMode(true);
    try {
      await onSwitchSessionMode(mode);
      setModePickerOpen(false);
    } catch (error) {
      setModeError(error instanceof Error ? error.message : 'Failed to switch conversation mode');
    } finally {
      setIsSwitchingMode(false);
    }
  };

  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      {modeError && (
        <div
          role="alert"
          className="px-3 py-1 text-xs"
          style={{ color: 'var(--vscode-errorForeground)' }}
        >
          Could not change conversation mode: {modeError}
        </div>
      )}
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
            <AgentPicker host={host} onSwitchAgent={onSwitchAgent} />
            <ModelPicker hasActiveSession={messages.length > 0} onSwitchModel={onSwitchModel} />
            <Popover.Root open={modePickerOpen} onOpenChange={setModePickerOpen}>
              <Popover.Trigger asChild>
                <button
                  type="button"
                  disabled={isSwitchingMode}
                  className="aui-mode-picker inline-flex min-w-0 max-w-full h-[22px] items-center gap-1 rounded-[var(--vscode-cornerRadius-small)] px-1.5 text-[11px] transition-colors hover:bg-[var(--vscode-toolbar-hoverBackground)]"
                  style={{ color: 'var(--vscode-icon-foreground)' }}
                  title={`Conversation mode: ${modeLabel}`}
                  aria-label="Select conversation mode"
                >
                  <Codicon
                    name={sessionMode === 'plan' ? 'checklist' : 'comment-discussion'}
                    className="shrink-0 text-xs"
                  />
                  <span className="min-w-0 truncate">
                    {isSwitchingMode ? 'Switching…' : modeLabel}
                  </span>
                  <Codicon name="chevron-down" className="shrink-0 text-[12px] opacity-60" />
                </button>
              </Popover.Trigger>
              <Popover.Portal>
                <Popover.Content
                  className="z-50 w-64 max-w-[calc(100vw-16px)] rounded-[var(--vscode-cornerRadius-medium)] border border-border bg-popover p-1 text-popover-foreground shadow-md outline-none"
                  sideOffset={4}
                  collisionPadding={8}
                  align="start"
                >
                  <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">
                    Conversation mode
                  </div>
                  {CONVERSATION_MODES.map(({ mode, label, description }) => (
                    <button
                      key={mode}
                      type="button"
                      aria-label={label}
                      aria-pressed={sessionMode === mode}
                      disabled={isSwitchingMode}
                      onClick={() => void selectMode(mode)}
                      className="flex w-full items-start gap-2 rounded-[var(--vscode-cornerRadius-medium)] px-2 py-1.5 text-left text-xs transition-colors hover:bg-accent disabled:opacity-50"
                    >
                      <Codicon
                        name="check"
                        className={`mt-0.5 shrink-0 text-[12px] ${sessionMode === mode ? 'opacity-100' : 'opacity-0'}`}
                      />
                      <span className="min-w-0">
                        <span className="block font-medium">{label}</span>
                        <span className="block text-muted-foreground">{description}</span>
                      </span>
                    </button>
                  ))}
                </Popover.Content>
              </Popover.Portal>
            </Popover.Root>
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
