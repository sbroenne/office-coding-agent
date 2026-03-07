import React from 'react';
import { MessageList } from '@/components/chat/MessageList';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import { McpPicker } from './McpPicker';
import type { ChatMessage } from '@/types';
import { useSettingsStore } from '@/stores/settingsStore';

interface ChatPanelProps {
  messages: ChatMessage[];
  isRunning: boolean;
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  onEnqueue?: (text: string) => void;
  queuedPrompts?: string[];
  onDequeue?: (index: number) => void;
  onOpenPanel?: (panel: string) => void;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({
  messages,
  isRunning,
  onSend,
  onCancel,
  onEnqueue,
  queuedPrompts,
  onDequeue,
  onOpenPanel,
}) => {
  const pluginPrompts = useSettingsStore(s => s.pluginPrompts);
  const setActiveAgent = useSettingsStore(s => s.setActiveAgent);

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
        slashCommands={pluginPrompts}
        onAgentSelect={setActiveAgent}
        leftToolbar={
          <>
            <AgentPicker onOpenPanel={onOpenPanel} />
            <ModelPicker />
          </>
        }
        rightToolbar={<McpPicker onOpenPanel={onOpenPanel} />}
      />
    </div>
  );
};
