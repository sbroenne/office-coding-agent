import React, { useMemo, useState } from 'react';
import * as Popover from '@radix-ui/react-popover';
import { History, ChevronDown } from 'lucide-react';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';

interface SessionHistoryPickerProps {
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  onRestoreSession: (sessionId: string) => void;
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
  sessions,
  activeSessionId,
  onRestoreSession,
}) => {
  const [open, setOpen] = useState(false);

  const ordered = useMemo(
    () => [...sessions].sort((a, b) => b.updatedAt - a.updatedAt),
    [sessions]
  );

  return (
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
              <button
                key={session.id}
                onClick={() => {
                  onRestoreSession(session.id);
                  setOpen(false);
                }}
                className="flex w-full items-start justify-between gap-2 rounded-md px-2 py-1.5 text-left text-sm transition-colors hover:bg-accent"
              >
                <div className="min-w-0 flex-1">
                  <div className={isActive ? 'font-medium text-foreground' : 'text-foreground'}>
                    {session.title}
                  </div>
                  <div className="text-xs text-muted-foreground">{session.host}</div>
                </div>
                <div className="shrink-0 text-[10px] text-muted-foreground">
                  {isActive ? 'active' : formatRelativeTime(session.updatedAt)}
                </div>
              </button>
            );
          })}
        </Popover.Content>
      </Popover.Portal>
    </Popover.Root>
  );
};
