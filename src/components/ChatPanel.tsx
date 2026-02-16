import React from 'react';
import { Thread } from '@/components/assistant-ui/thread';
import { AgentPicker } from './AgentPicker';
import { ModelPicker } from './ModelPicker';

export interface ChatPanelProps {
  isConfigured: boolean;
  onOpenSettings?: () => void;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({ isConfigured, onOpenSettings }) => {
  return (
    <div className="flex flex-1 flex-col overflow-hidden">
      {/* Not-configured warning */}
      {!isConfigured && (
        <div className="mx-3 mt-2 rounded-md border border-yellow-300 bg-yellow-50 px-3 py-2 text-sm text-yellow-800 dark:border-yellow-700 dark:bg-yellow-900/30 dark:text-yellow-200">
          No model configured.{' '}
          {onOpenSettings ? (
            <button onClick={onOpenSettings} className="font-medium underline hover:no-underline">
              Open Settings
            </button>
          ) : (
            'Open Settings to get started.'
          )}
        </div>
      )}

      {/* Chat thread */}
      <Thread />

      {/* Input toolbar: Agent & Model pickers */}
      <div className="flex items-center gap-1 border-t border-border bg-background px-3 py-1">
        <AgentPicker />
        <ModelPicker onOpenSettings={onOpenSettings} />
      </div>
    </div>
  );
};
