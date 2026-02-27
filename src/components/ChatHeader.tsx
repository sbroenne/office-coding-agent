import React from 'react';
import { RotateCcw, Shield } from 'lucide-react';
import { SessionHistoryPicker } from './SessionHistoryPicker';
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
        <button
          onClick={() => onOpenPanel?.('permissions')}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="Permissions"
          title="Permissions"
        >
          <Shield className="size-4" />
        </button>
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
