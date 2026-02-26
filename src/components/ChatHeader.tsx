import React, { useState } from 'react';
import { RotateCcw, ServerIcon, BrainCircuit, ChevronDown } from 'lucide-react';
import { SkillPicker } from './SkillPicker';
import { SessionHistoryPicker } from './SessionHistoryPicker';
import { McpManagerDialog } from './McpManagerDialog';
import { PermissionManagerDialog } from './PermissionManagerDialog';
import { useSettingsStore } from '@/stores';
import type { SessionHistoryItem } from '@/stores/sessionHistoryStore';
import type { OfficeHostApp } from '@/services/office/host';

export interface ChatHeaderProps {
  host: OfficeHostApp;
  onClearMessages: () => void;
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  onRestoreSession: (sessionId: string) => void;
  onDeleteSession: (sessionId: string) => void;
}

export const ChatHeader: React.FC<ChatHeaderProps> = ({
  host,
  onClearMessages,
  sessions,
  activeSessionId,
  onRestoreSession,
  onDeleteSession,
}) => {
  const [mcpOpen, setMcpOpen] = useState(false);
  const workiqEnabled = useSettingsStore(s => s.workiqEnabled);
  const toggleWorkiq = useSettingsStore(s => s.toggleWorkiq);
  const workiqModel = useSettingsStore(s => s.workiqModel);
  const setWorkiqModel = useSettingsStore(s => s.setWorkiqModel);
  const availableModels = useSettingsStore(s => s.availableModels);
  return (
    <div className="flex items-center justify-between border-b border-border bg-background px-3 py-1.5">
      <div className="flex items-center gap-2 min-w-0">
        <SkillPicker />
        <SessionHistoryPicker
          host={host}
          sessions={sessions}
          activeSessionId={activeSessionId}
          onRestoreSession={onRestoreSession}
          onDeleteSession={onDeleteSession}
        />
      </div>

      <div className="flex items-center gap-0.5">
        <PermissionManagerDialog />
        <div className="flex items-center">
          <button
            onClick={toggleWorkiq}
            className={`inline-flex h-8 items-center gap-1.5 rounded-l-md px-2 text-xs font-medium transition-colors ${
              workiqEnabled
                ? 'bg-emerald-100 text-emerald-700 hover:bg-emerald-200 dark:bg-emerald-900/40 dark:text-emerald-300 dark:hover:bg-emerald-900/60'
                : 'bg-muted/50 text-muted-foreground line-through hover:bg-accent hover:text-accent-foreground'
            }`}
            aria-label={`WorkIQ: ${workiqEnabled ? 'on' : 'off'}`}
            aria-pressed={workiqEnabled}
            title={
              workiqEnabled
                ? 'WorkIQ enabled — click to disable'
                : 'Enable WorkIQ (Microsoft 365 data)'
            }
          >
            <span className="relative">
              <BrainCircuit className="size-4" />
              {workiqEnabled && (
                <span className="absolute -right-0.5 -top-0.5 size-2 rounded-full bg-emerald-500" />
              )}
            </span>
            <span>WorkIQ</span>
          </button>
          {workiqEnabled && availableModels && availableModels.length > 0 && (
            <div className="relative">
              <select
                value={workiqModel ?? ''}
                onChange={e => setWorkiqModel(e.target.value || null)}
                className="h-8 max-w-[80px] appearance-none rounded-r-md border-l border-emerald-300 bg-emerald-50 pl-3 pr-4 text-[9px] font-medium text-emerald-800 hover:bg-emerald-100 focus:outline-none dark:border-emerald-700 dark:bg-emerald-950 dark:text-emerald-200"
                title="WorkIQ model (defaults to main model)"
                aria-label="WorkIQ model"
              >
                <option value="">Model</option>
                {availableModels.map(m => (
                  <option key={m.id} value={m.id}>
                    {m.name}
                  </option>
                ))}
              </select>
              <ChevronDown className="pointer-events-none absolute right-1 top-1/2 size-3 -translate-y-1/2 text-emerald-700 dark:text-emerald-300" />
            </div>
          )}
        </div>
        <button
          onClick={() => setMcpOpen(true)}
          className="inline-flex h-8 w-8 items-center justify-center rounded-md text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors"
          aria-label="MCP Servers"
          title="MCP Servers"
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
      </div>

      <McpManagerDialog open={mcpOpen} onOpenChange={setMcpOpen} />
    </div>
  );
};
