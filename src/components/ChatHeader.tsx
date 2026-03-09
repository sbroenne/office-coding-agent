import React from 'react';
import { SessionHistoryPicker } from './SessionHistoryPicker';
import { SkillPicker } from './SkillPicker';
import { Codicon } from '@/components/Codicon';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';
import type { OfficeHostApp } from '@/services/office/host';

export interface ChatHeaderProps {
  host: OfficeHostApp;
  onClearMessages: () => void;
  onCompactSession?: () => void;
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  onRestoreSession: (sessionId: string) => void;
  onDeleteSession: (sessionId: string) => void;
  onOpenPanel?: (panel: string) => void;
}

export const ChatHeader: React.FC<ChatHeaderProps> = ({
  host: _host,
  onClearMessages,
  onCompactSession,
  sessions,
  activeSessionId,
  onRestoreSession,
  onDeleteSession,
  onOpenPanel,
}) => {
  return (
    <div className="flex h-[35px] items-center justify-between border-b border-border bg-background px-2">
      <div className="flex items-center gap-1 min-w-0">
        <SessionHistoryPicker
          sessions={sessions}
          activeSessionId={activeSessionId}
          onRestoreSession={onRestoreSession}
          onDeleteSession={onDeleteSession}
          onOpenPanel={onOpenPanel}
        />
      </div>

      <div className="flex items-center gap-1">
        <SkillPicker />
        <button
          onClick={() => onOpenPanel?.('permissions')}
          className="inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
          style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
          aria-label="Permissions"
          title="Permissions"
        >
          <Codicon name="shield" className="text-[14px]" />
        </button>
        <button
          onClick={onCompactSession}
          className="inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
          style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
          aria-label="Compact conversation"
          title="Compact conversation"
        >
          <Codicon name="fold" className="text-[14px]" />
        </button>
        <button
          onClick={onClearMessages}
          className="inline-flex items-center justify-center rounded-[var(--vscode-cornerRadius-small)] transition-colors hover:bg-accent"
          style={{ width: 22, height: 22, color: 'var(--vscode-icon-foreground)' }}
          aria-label="New conversation"
          title="New conversation"
        >
          <Codicon name="comment-add" className="text-[14px]" />
        </button>
      </div>
    </div>
  );
};
