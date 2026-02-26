import React from 'react';
import { Thread } from '@/components/assistant-ui/thread';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';

interface ChatPanelProps {
  onOpenPanel?: (panel: string) => void;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({ onOpenPanel }) => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      {/* Chat thread */}
      <Thread />

      {/* Input toolbar: Agent & Model pickers */}
      <div className="flex items-center gap-1 border-t border-border bg-background px-3 py-1">
        <AgentPicker onOpenPanel={onOpenPanel} />
        <ModelPicker />
      </div>
    </div>
  );
};
