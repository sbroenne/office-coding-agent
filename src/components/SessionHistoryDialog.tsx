import React, { useMemo } from 'react';
import { Trash2 } from 'lucide-react';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';
import type { OfficeHostApp } from '@/services/office/host';

interface SessionHistoryPanelProps {
  host: OfficeHostApp;
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  onRestoreSession: (sessionId: string) => void;
  onDeleteSession: (sessionId: string) => void;
}

function formatDate(timestamp: number): string {
  return new Date(timestamp).toLocaleString();
}

export const SessionHistoryPanel: React.FC<SessionHistoryPanelProps> = ({
  host,
  sessions,
  activeSessionId,
  onRestoreSession,
  onDeleteSession,
}) => {
  const hostSessions = useMemo(
    () => sessions.filter(s => s.host === host).sort((a, b) => b.updatedAt - a.updatedAt),
    [host, sessions]
  );

  return (
    <div className="space-y-3 p-3">
      <p className="text-xs text-muted-foreground">
        {hostSessions.length} saved conversation{hostSessions.length === 1 ? '' : 's'} for {host}.
      </p>

      {hostSessions.length === 0 ? (
        <div className="text-sm text-muted-foreground">No saved conversations yet.</div>
      ) : (
        <div className="space-y-2">
          {hostSessions.map(session => {
            const isActive = session.id === activeSessionId;
            return (
              <div key={session.id} className="rounded-md border border-border p-2 text-sm">
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0 flex-1">
                    <div className={isActive ? 'font-medium' : ''}>{session.title}</div>
                    <div className="text-xs text-muted-foreground">
                      Updated {formatDate(session.updatedAt)}
                    </div>
                  </div>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => onRestoreSession(session.id)}
                      className="rounded-md border border-border px-2 py-1 text-xs hover:bg-accent"
                    >
                      {isActive ? 'Active' : 'Restore'}
                    </button>
                    <button
                      type="button"
                      onClick={() => onDeleteSession(session.id)}
                      className="rounded-md p-1 text-muted-foreground hover:bg-accent hover:text-foreground"
                      aria-label="Delete session"
                      title="Delete session"
                    >
                      <Trash2 className="size-3" />
                    </button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
