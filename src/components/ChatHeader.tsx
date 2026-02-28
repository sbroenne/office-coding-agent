import React from 'react';
import { SessionHistoryPicker } from './SessionHistoryPicker';
import { SkillPicker } from './SkillPicker';
import { Codicon } from '@/components/Codicon';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';
import type { OfficeHostApp } from '@/services/office/host';

export interface ChatHeaderProps {
  host: OfficeHostApp;
  onClearMessages: () => void;
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  onRestoreSession: (sessionId: string) => void;
  onDeleteSession: (sessionId: string) => void;
  onOpenPanel?: (panel: string) => void;
}

export const ChatHeader: React.FC<ChatHeaderProps> = ({
  host: _host,
  onClearMessages,
  sessions,
  activeSessionId,
  onRestoreSession,
  onDeleteSession,
  onOpenPanel,
}) => {
  return (
    <div className="flex items-center justify-between border-b border-border bg-background px-3 py-1.5">
      <div className="flex items-center gap-2 min-w-0">
        <SessionHistoryPicker
          sessions={sessions}
          activeSessionId={activeSessionId}
          onRestoreSession={onRestoreSession}
          onDeleteSession={onDeleteSession}
          onOpenPanel={onOpenPanel}
        />
      </div>

      <div className="flex items-center gap-0.5">
        <SkillPicker onOpenPanel={onOpenPanel} />
        <button
          onClick={() => onOpenPanel?.('permissions')}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="Permissions"
          title="Permissions"
        >
          <Codicon name="shield" className="text-base" />
        </button>
        <button
          onClick={onClearMessages}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="New conversation"
          title="New conversation"
        >
          <Codicon name="comment-add" className="text-base" />
        </button>
      </div>
    </div>
  );
};
