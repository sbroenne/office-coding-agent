import React from 'react';
import { MessageList } from '@/components/chat/MessageList';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';
import type { ChatMessage } from '@/types';

interface ChatPanelProps {
  messages: ChatMessage[];
  isRunning: boolean;
  onSend: (text: string) => void | Promise<void>;
  onCancel: () => void;
  onEnqueue?: (text: string) => void;
  queuedCount?: number;
  onOpenPanel?: (panel: string) => void;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({
  messages,
  isRunning,
  onSend,
  onCancel,
  onEnqueue,
  queuedCount,
  onOpenPanel,
}) => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      <MessageList
        messages={messages}
        isRunning={isRunning}
        onSend={onSend}
        onCancel={onCancel}
        onEnqueue={onEnqueue}
        queuedCount={queuedCount}
        onFeedback={() => {
          /* TODO */
        }}
        onRegenerate={() => {
          /* TODO */
        }}
        leftToolbar={
          <>
            <AgentPicker onOpenPanel={onOpenPanel} />
            <ModelPicker />
          </>
        }
        rightToolbar={null}
      />
    </div>
  );
};
