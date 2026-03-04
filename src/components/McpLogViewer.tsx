import React, { useEffect, useRef } from 'react';
import { Codicon } from '@/components/Codicon';
import { Button } from '@/components/ui/button';
import { useMcpStatusStore } from '@/stores';

interface McpLogViewerProps {
  serverName: string | null;
}

export const McpLogViewer: React.FC<McpLogViewerProps> = ({ serverName }) => {
  const servers = useMcpStatusStore(s => s.servers);
  const clearLogs = useMcpStatusStore(s => s.clearLogs);
  const scrollRef = useRef<HTMLDivElement>(null);
  const autoScrollRef = useRef(true);

  const logs = serverName ? (servers[serverName]?.logs ?? []) : [];

  useEffect(() => {
    if (autoScrollRef.current && scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [logs.length]);

  const handleScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 40;
    autoScrollRef.current = atBottom;
  };

  const handleCopy = async () => {
    const text = logs.map(l => `${l.timestamp} [${l.level}] ${l.message}`).join('\n');
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // clipboard unavailable in some contexts
    }
  };

  if (!serverName) {
    return (
      <div className="rounded-md border border-border bg-muted/30 p-3 text-center text-xs text-muted-foreground">
        Select a server to view its output
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-1">
      <div className="flex items-center justify-between">
        <span className="text-xs font-medium text-muted-foreground">Output: {serverName}</span>
        <div className="flex items-center gap-1">
          <Button
            variant="ghost"
            size="sm"
            className="h-6 px-1.5 text-xs"
            onClick={() => void handleCopy()}
            title="Copy logs"
            disabled={logs.length === 0}
          >
            <Codicon name="copy" className="text-xs" />
          </Button>
          <Button
            variant="ghost"
            size="sm"
            className="h-6 px-1.5 text-xs"
            onClick={() => clearLogs(serverName)}
            title="Clear logs"
            disabled={logs.length === 0}
          >
            <Codicon name="trash" className="text-xs" />
          </Button>
        </div>
      </div>
      <div
        ref={scrollRef}
        onScroll={handleScroll}
        className="h-[160px] overflow-y-auto rounded-md border border-border bg-black/90 p-2 font-mono text-[10px] leading-4"
      >
        {logs.length === 0 ? (
          <span className="text-muted-foreground">
            No output yet — logs appear when the server connects
          </span>
        ) : (
          logs.map((entry, i) => (
            <div
              key={i}
              className={
                entry.level === 'error'
                  ? 'text-[var(--vscode-errorForeground)]'
                  : entry.level === 'warn'
                    ? 'text-[var(--vscode-descriptionForeground)]'
                    : 'text-foreground'
              }
            >
              <span className="text-muted-foreground">
                {entry.timestamp.split('T')[1]?.split('.')[0] ?? entry.timestamp}
              </span>{' '}
              {entry.message}
            </div>
          ))
        )}
      </div>
    </div>
  );
};
