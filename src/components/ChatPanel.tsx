import React from 'react';
import { Thread } from '@/components/assistant-ui/thread';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';

export const ChatPanel: React.FC = () => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      {/* Chat thread */}
      <Thread />

      {/* Input toolbar: Agent & Model pickers */}
      <div className="flex items-center gap-1 border-t border-border bg-background px-3 py-1">
        <AgentPicker />
        <ModelPicker />
      </div>
    </div>
  );
};
