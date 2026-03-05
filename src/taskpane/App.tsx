import React, { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import { Loader2, RefreshCw } from 'lucide-react';
import { ChatHeader } from '@/components/ChatHeader';
import { ChatPanel } from '@/components/ChatPanel';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';
import { SlidePanel } from '@/components/SlidePanel';
import { McpManagerPanel } from '@/components/McpManagerDialog';
import { AgentManagerPanel } from '@/components/AgentManagerDialog';
import { SkillManagerPanel } from '@/components/SkillManagerDialog';
import { SessionHistoryPanel } from '@/components/SessionHistoryDialog';
import { PermissionManagerPanel } from '@/components/PermissionManagerDialog';
import { useSettingsStore } from '@/stores';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { ChatActionsContext } from '@/contexts/ChatActionsContext';
import { detectOfficeHost } from '@/services/office/host';
import type { OfficeHostApp } from '@/services/office/host';

const ConnectingBanner: React.FC = () => (
  <div className="flex items-center gap-2 border-b border-border bg-muted/50 px-3 py-2 text-sm text-muted-foreground">
    <Loader2 className="size-3.5 animate-spin shrink-0" />
    <span>Connecting to Copilot...</span>
  </div>
);

const SessionErrorBanner: React.FC<{ error: Error; onRetry: () => void }> = ({
  error,
  onRetry,
}) => (
  <div className="flex items-center gap-2 border-b border-destructive bg-destructive/10 px-3 py-2 text-sm text-destructive dark:text-red-200">
    <span className="min-w-0 flex-1 truncate" title={error.message}>
      Connection failed: {error.message}
    </span>
    <button
      onClick={onRetry}
      className="flex items-center gap-1 shrink-0 rounded-md border border-destructive/30 px-2 py-0.5 text-xs font-medium hover:bg-destructive/20 transition-colors"
    >
      <RefreshCw className="size-3" />
      Retry
    </button>
  </div>
);

const PermissionBanner: React.FC<{
  kind: string;
  detail: string;
  onApprove: () => void;
  onDeny: () => void;
  onAlwaysAllow: () => void;
}> = ({ kind, detail, onApprove, onDeny, onAlwaysAllow }) => (
  <div className="flex items-center gap-2 border-b border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-900 dark:border-amber-700 dark:bg-amber-950/30 dark:text-amber-200">
    <div className="min-w-0 flex-1">
      <div className="truncate font-medium" title={kind}>
        Permission requested: {kind}
      </div>
      <div className="truncate text-xs" title={detail}>
        {detail}
      </div>
    </div>
    <div className="flex items-center gap-1">
      <button
        onClick={onDeny}
        className="shrink-0 rounded-md border border-amber-500/40 px-2 py-0.5 text-xs font-medium hover:bg-amber-200/60 dark:hover:bg-amber-900/40"
      >
        Deny
      </button>
      <button
        onClick={onApprove}
        className="shrink-0 rounded-md border border-amber-600/40 bg-amber-600 px-2 py-0.5 text-xs font-medium text-white hover:bg-amber-700"
      >
        Allow
      </button>
      <button
        onClick={onAlwaysAllow}
        className="shrink-0 rounded-md border border-amber-600/40 px-2 py-0.5 text-xs font-medium hover:bg-amber-200/60 dark:hover:bg-amber-900/40"
      >
        Always allow
      </button>
    </div>
  </div>
);

const PANEL_TITLES: Record<string, { title: string; description?: string }> = {
  mcp: { title: 'MCP Servers', description: 'Manage servers that provide additional AI tools' },
  agents: { title: 'Manage Agents', description: 'Import and manage custom agents' },
  skills: { title: 'Manage Skills', description: 'Import and manage custom skills' },
  history: { title: 'Session History' },
  permissions: { title: 'Permissions', description: 'Manage auto-approval and saved rules' },
};

const ReadyAssistant: React.FC<{ host: OfficeHostApp }> = ({ host }) => {
  const {
    messages,
    isRunning,
    send,
    cancel,
    sessionError,
    isConnecting,
    clearMessages,
    compactSession,
    restoreSession,
    deleteSession,
    sessions,
    activeSessionId,
    pendingPermission,
    approvePermission,
    denyPermission,
    allowPermissionAlways,
    enqueue,
    queuedPrompts,
  } = useOfficeChat(host);

  const [activePanel, setActivePanel] = useState<string | null>(null);
  const closePanel = useCallback(() => setActivePanel(null), []);

  const permissionDetail = pendingPermission
    ? (pendingPermission.request.path ??
      pendingPermission.request.fileName ??
      pendingPermission.request.fullCommandText ??
      pendingPermission.request.intention ??
      'User approval required')
    : '';

  return (
    <ChatActionsContext.Provider value={{ send, enqueue }}>
      <div className="relative flex h-screen flex-col overflow-hidden bg-background text-foreground">
        {/* Main chat view */}
        <div
          className={`flex h-full flex-col will-change-transform transition-transform duration-300 ease-in-out ${
            activePanel ? '-translate-x-full' : 'translate-x-0'
          }`}
          aria-hidden={!!activePanel}
          inert={activePanel ? ('' as unknown as boolean) : undefined}
        >
          <ChatHeader
            host={host}
            onClearMessages={clearMessages}
            onCompactSession={() => void compactSession()}
            sessions={sessions}
            activeSessionId={activeSessionId}
            onRestoreSession={restoreSession}
            onDeleteSession={deleteSession}
            onOpenPanel={setActivePanel}
          />
          {isConnecting && !sessionError && <ConnectingBanner />}
          {sessionError && <SessionErrorBanner error={sessionError} onRetry={clearMessages} />}
          {pendingPermission && (
            <PermissionBanner
              kind={pendingPermission.request.kind}
              detail={permissionDetail}
              onApprove={approvePermission}
              onDeny={denyPermission}
              onAlwaysAllow={allowPermissionAlways}
            />
          )}
          <ChatErrorBoundary>
            <ChatPanel
              messages={messages}
              isRunning={isRunning}
              onSend={send}
              onCancel={cancel}
              onEnqueue={enqueue}
              queuedCount={queuedPrompts.length}
              onOpenPanel={setActivePanel}
            />
          </ChatErrorBoundary>
        </div>

        {/* Slide panels */}
        {Object.entries(PANEL_TITLES).map(([key, { title, description }]) => (
          <SlidePanel
            key={key}
            open={activePanel === key}
            onClose={closePanel}
            title={title}
            description={description}
          >
            {key === 'mcp' && <McpManagerPanel />}
            {key === 'agents' && <AgentManagerPanel />}
            {key === 'skills' && <SkillManagerPanel />}
            {key === 'history' && (
              <SessionHistoryPanel
                host={host}
                sessions={sessions}
                activeSessionId={activeSessionId}
                onRestoreSession={restoreSession}
                onDeleteSession={deleteSession}
              />
            )}
            {key === 'permissions' && <PermissionManagerPanel />}
          </SlidePanel>
        ))}
      </div>
    </ChatActionsContext.Provider>
  );
};

export const App: React.FC = () => {
  // Wait for Zustand persist to finish hydrating from async storage.
  const hasHydrated = useSyncExternalStore(useSettingsStore.persist.onFinishHydration, () =>
    useSettingsStore.persist.hasHydrated()
  );

  // Detect system theme preference, reacting to OS changes
  const prefersDark = useSyncExternalStore(
    onStoreChange => {
      if (typeof window === 'undefined') return () => undefined;
      const mql = window.matchMedia('(prefers-color-scheme: dark)');
      mql.addEventListener('change', onStoreChange);
      return () => mql.removeEventListener('change', onStoreChange);
    },
    () => typeof window !== 'undefined' && window.matchMedia('(prefers-color-scheme: dark)').matches
  );

  // Sync .dark class on <html> so Tailwind dark: variants work
  useEffect(() => {
    document.documentElement.classList.toggle('dark', prefersDark);
  }, [prefersDark]);

  if (!hasHydrated) {
    return (
      <div className="flex h-screen flex-col items-center justify-center gap-3 bg-background text-foreground">
        <Loader2 className="size-8 animate-spin text-muted-foreground" />
        <p className="text-sm text-muted-foreground">Initializing...</p>
      </div>
    );
  }

  const host = detectOfficeHost();
  return <ReadyAssistant host={host} />;
};
