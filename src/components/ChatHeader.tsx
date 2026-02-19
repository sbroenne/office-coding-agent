import React, { useState } from 'react';
import { RotateCcw, ServerIcon } from 'lucide-react';
import { SkillPicker } from './SkillPicker';
import { SettingsDialog } from './SettingsDialog';
import { McpManagerDialog } from './McpManagerDialog';

export interface ChatHeaderProps {
  onClearMessages: () => void;
  settingsOpen: boolean;
  onSettingsOpenChange: (open: boolean) => void;
}

export const ChatHeader: React.FC<ChatHeaderProps> = ({
  onClearMessages,
  settingsOpen,
  onSettingsOpenChange,
}) => {
  const [mcpOpen, setMcpOpen] = useState(false);

  return (
    <div className="flex items-center justify-between border-b border-border bg-background px-3 py-1.5">
      <div className="flex items-center gap-2 min-w-0">
        <span className="font-semibold text-sm whitespace-nowrap text-foreground">AI Chat</span>
        <SkillPicker />
      </div>

      <div className="flex items-center gap-0.5">
        <button
          onClick={() => setMcpOpen(true)}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="MCP servers"
          title="MCP servers"
        >
          <ServerIcon className="size-4" />
        </button>
        <button
          onClick={onClearMessages}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="New conversation"
          title="New conversation"
        >
          <RotateCcw className="size-4" />
        </button>
        <div className="mx-1 h-4 w-px bg-border" />
        <SettingsDialog open={settingsOpen} onOpenChange={onSettingsOpenChange} />
        <McpManagerDialog open={mcpOpen} onOpenChange={setMcpOpen} />
      </div>
    </div>
  );
};
