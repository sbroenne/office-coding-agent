import React, { useMemo, useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { History, ChevronDown, Trash2 } from 'lucide-react';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';
import type { OfficeHostApp } from '@/services/office/host';
import { SessionHistoryDialog } from './SessionHistoryDialog';

interface SessionHistoryPickerProps {
  host: OfficeHostApp;
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  onRestoreSession: (sessionId: string) => void;
  onDeleteSession: (sessionId: string) => void;
}

function formatRelativeTime(updatedAt: number): string {
  const diffMs = Date.now() - updatedAt;
  const minutes = Math.floor(diffMs / 60_000);
  if (minutes < 1) return 'just now';
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  const days = Math.floor(hours / 24);
  return `${days}d ago`;
}

export const SessionHistoryPicker: React.FC<SessionHistoryPickerProps> = ({
  host,
  sessions,
  activeSessionId,
  onRestoreSession,
  onDeleteSession,
}) => {
  const [open, setOpen] = useState(false);
  const [manageOpen, setManageOpen] = useState(false);

  const ordered = useMemo(
    () => [...sessions].sort((a, b) => b.updatedAt - a.updatedAt),
    [sessions]
  );

  return (
    <>
      <Popover.Root open={open} onOpenChange={setOpen}>
        <Popover.Trigger asChild>
          <button
            className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
            aria-label="Session history"
            title="Session history"
          >
            <History className="size-4" />
            <ChevronDown className="size-3 opacity-60" />
          </button>
        </Popover.Trigger>

        <Popover.Portal>
          <Popover.Content
            className="z-50 w-72 max-h-80 overflow-y-auto rounded-lg border border-border bg-popover p-1 shadow-md outline-none"
            sideOffset={4}
            align="start"
          >
            <div className="px-2 py-1.5 text-xs font-medium text-muted-foreground">
              Session history
            </div>

            {ordered.length === 0 && (
              <div className="px-2 py-2 text-xs text-muted-foreground">
                No previous conversations yet.
              </div>
            )}

            {ordered.map(session => {
              const isActive = session.id === activeSessionId;
              return (
                <div
                  key={session.id}
                  className="flex w-full items-start justify-between gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent"
                >
                  <div className="min-w-0 flex-1">
                    <button
                      type="button"
                      onClick={() => {
                        onRestoreSession(session.id);
                        setOpen(false);
                      }}
                      className={isActive ? 'font-medium text-foreground' : 'text-foreground'}
                    >
                      {session.title}
                    </button>
                    <div className="text-xs text-muted-foreground">{session.host}</div>
                  </div>
                  <div className="flex items-center gap-1">
                    <div className="shrink-0 text-[10px] text-muted-foreground">
                      {isActive ? 'active' : formatRelativeTime(session.updatedAt)}
                    </div>
                    <button
                      type="button"
                      onClick={() => onDeleteSession(session.id)}
                      className="rounded p-1 text-muted-foreground hover:bg-accent hover:text-foreground"
                      aria-label="Delete session"
                      title="Delete session"
                    >
                      <Trash2 className="size-3" />
                    </button>
                  </div>
                </div>
              );
            })}

            {ordered.length > 0 && (
              <div className="mt-1 border-t border-border pt-1">
                <button
                  type="button"
                  onClick={() => {
                    setOpen(false);
                    setManageOpen(true);
                  }}
                  className="flex w-full items-center justify-between rounded-md px-2 py-1.5 text-left text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
                >
                  <span>Manage history…</span>
                  <span>{ordered.length} total</span>
                </button>
              </div>
            )}
          </Popover.Content>
        </Popover.Portal>
      </Popover.Root>

      <SessionHistoryDialog
        open={manageOpen}
        onOpenChange={setManageOpen}
        host={host}
        sessions={sessions}
        activeSessionId={activeSessionId}
        onRestoreSession={onRestoreSession}
        onDeleteSession={onDeleteSession}
      />
    </>
  );
};
