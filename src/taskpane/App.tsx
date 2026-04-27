import React, { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import { ChatHeader } from '@/components/ChatHeader';
import { ChatPanel } from '@/components/ChatPanel';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';
import { SlidePanel } from '@/components/SlidePanel';
import { SessionHistoryPanel } from '@/components/SessionHistoryDialog';
import { PermissionManagerPanel } from '@/components/PermissionManagerDialog';
import { Codicon } from '@/components/Codicon';
import { useSettingsStore } from '@/stores';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { ChatActionsContext } from '@/contexts/ChatActionsContext';
import { detectOfficeHost } from '@/services/office/host';
import type { OfficeHostApp } from '@/services/office/host';

const bannerBaseStyle: React.CSSProperties = {
  borderBottomColor: 'var(--vscode-widget-border)',
  backgroundColor: 'var(--vscode-editorWidget-background)',
  color: 'var(--vscode-descriptionForeground)',
};

const buttonBaseClassName =
  'shrink-0 rounded-[var(--vscode-cornerRadius-medium)] border px-2 py-0.5 text-xs font-medium transition-colors focus-visible:outline focus-visible:outline-1 focus-visible:outline-[var(--vscode-focusBorder)] focus-visible:outline-offset-0';

const ConnectingBanner: React.FC = () => (
  <div className="flex items-center gap-2 border-b px-3 py-2 text-sm" style={bannerBaseStyle}>
    <Codicon name="loading" className="shrink-0 animate-spin text-[14px]" />
    <span>Connecting to Copilot...</span>
  </div>
);

const SessionErrorBanner: React.FC<{ error: Error; onRetry: () => void }> = ({
  error,
  onRetry,
}) => (
  <div
    className="flex items-center gap-2 border-b px-3 py-2 text-sm"
    style={{
      ...bannerBaseStyle,
      borderBottomColor: 'var(--vscode-errorForeground)',
      color: 'var(--vscode-errorForeground)',
    }}
  >
    <span className="min-w-0 flex-1 truncate" title={error.message}>
      Connection failed: {error.message}
    </span>
    <button
      onClick={onRetry}
      className={buttonBaseClassName}
      style={{
        borderColor: 'var(--vscode-widget-border)',
        color: 'var(--vscode-foreground)',
        backgroundColor: 'transparent',
      }}
      onMouseEnter={event => {
        event.currentTarget.style.backgroundColor = 'var(--vscode-toolbar-hoverBackground)';
      }}
      onMouseLeave={event => {
        event.currentTarget.style.backgroundColor = 'transparent';
      }}
    >
      <Codicon name="refresh" className="text-[12px]" />
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
  <div
    className="flex items-center gap-2 border-b px-3 py-2 text-sm"
    style={{
      ...bannerBaseStyle,
      color: 'var(--vscode-foreground)',
    }}
  >
    <Codicon name="warning" className="shrink-0 text-[14px]" />
    <div className="min-w-0 flex-1">
      <div className="truncate font-medium" title={kind}>
        Permission requested: {kind}
      </div>
      <div
        className="truncate text-xs"
        title={detail}
        style={{ color: 'var(--vscode-descriptionForeground)' }}
      >
        {detail}
      </div>
    </div>
    <div className="flex items-center gap-1">
      <button
        onClick={onDeny}
        className={buttonBaseClassName}
        style={{
          borderColor: 'var(--vscode-widget-border)',
          backgroundColor: 'var(--vscode-button-secondaryBackground)',
          color: 'var(--vscode-button-secondaryForeground)',
        }}
        onMouseEnter={event => {
          event.currentTarget.style.backgroundColor =
            'var(--vscode-button-secondaryHoverBackground)';
        }}
        onMouseLeave={event => {
          event.currentTarget.style.backgroundColor = 'var(--vscode-button-secondaryBackground)';
        }}
      >
        Deny
      </button>
      <button
        onClick={onApprove}
        className={buttonBaseClassName}
        style={{
          borderColor: 'var(--vscode-button-background)',
          backgroundColor: 'var(--vscode-button-background)',
          color: 'var(--vscode-button-foreground)',
        }}
        onMouseEnter={event => {
          event.currentTarget.style.backgroundColor = 'var(--vscode-button-hoverBackground)';
        }}
        onMouseLeave={event => {
          event.currentTarget.style.backgroundColor = 'var(--vscode-button-background)';
        }}
      >
        Allow
      </button>
      <button
        onClick={onAlwaysAllow}
        className={buttonBaseClassName}
        style={{
          borderColor: 'var(--vscode-widget-border)',
          backgroundColor: 'transparent',
          color: 'var(--vscode-foreground)',
        }}
        onMouseEnter={event => {
          event.currentTarget.style.backgroundColor = 'var(--vscode-toolbar-hoverBackground)';
        }}
        onMouseLeave={event => {
          event.currentTarget.style.backgroundColor = 'transparent';
        }}
      >
        Always allow
      </button>
    </div>
  </div>
);

const PANEL_TITLES: Record<string, { title: string; description?: string }> = {
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
    switchModel,
    restoreSession,
    deleteSession,
    sessions,
    activeSessionId,
    pendingPermission,
    approvePermission,
    denyPermission,
    allowPermissionAlways,
    initiateMcpOAuth,
    enqueue,
    queuedPrompts,
    dequeue,
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
              onSwitchModel={switchModel}
              onInitiateMcpOAuth={initiateMcpOAuth}
              onEnqueue={enqueue}
              queuedPrompts={queuedPrompts}
              onDequeue={dequeue}
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
        <Codicon name="loading" className="animate-spin text-[32px]" />
        <p className="text-sm text-muted-foreground">Initializing...</p>
      </div>
    );
  }

  const host = detectOfficeHost();
  return <ReadyAssistant host={host} />;
};
