import React, { useEffect, useSyncExternalStore } from 'react';
import { Loader2, RefreshCw } from 'lucide-react';
import { AssistantRuntimeProvider } from '@assistant-ui/react';
import { ChatHeader } from '@/components/ChatHeader';
import { ChatPanel } from '@/components/ChatPanel';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';
import { useSettingsStore } from '@/stores';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { ThinkingContext } from '@/contexts/ThinkingContext';
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

const ReadyAssistant: React.FC<{ host: OfficeHostApp }> = ({ host }) => {
  const {
    runtime,
    sessionError,
    isConnecting,
    clearMessages,
    restoreSession,
    deleteSession,
    sessions,
    activeSessionId,
    pendingPermission,
    approvePermission,
    denyPermission,
    allowPermissionAlways,
    thinkingText,
  } = useOfficeChat(host);

  const permissionDetail = pendingPermission
    ? (pendingPermission.request.path ??
      pendingPermission.request.fileName ??
      pendingPermission.request.fullCommandText ??
      pendingPermission.request.intention ??
      'User approval required')
    : '';

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <ThinkingContext.Provider value={thinkingText}>
        <div className="flex h-screen flex-col overflow-hidden bg-background text-foreground">
          <ChatHeader
            host={host}
            onClearMessages={clearMessages}
            sessions={sessions}
            activeSessionId={activeSessionId}
            onRestoreSession={restoreSession}
            onDeleteSession={deleteSession}
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
            <ChatPanel />
          </ChatErrorBoundary>
        </div>
      </ThinkingContext.Provider>
    </AssistantRuntimeProvider>
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
