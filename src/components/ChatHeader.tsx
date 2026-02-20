import React from 'react';
import { RotateCcw } from 'lucide-react';
import { SkillPicker } from './SkillPicker';

export interface ChatHeaderProps {
  onClearMessages: () => void;
}

export const ChatHeader: React.FC<ChatHeaderProps> = ({ onClearMessages }) => {
  return (
    <div className="flex items-center justify-between border-b border-border bg-background px-3 py-1.5">
      <div className="flex items-center gap-2 min-w-0">
        <SkillPicker />
      </div>

      <div className="flex items-center gap-0.5">
        <button
          onClick={onClearMessages}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="New conversation"
          title="New conversation"
        >
          <RotateCcw className="size-4" />
        </button>
      </div>
    </div>
  );
};
