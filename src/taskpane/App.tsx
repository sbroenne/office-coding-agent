import React, { useEffect, useSyncExternalStore } from 'react';
import { Loader2, RefreshCw } from 'lucide-react';
import { AssistantRuntimeProvider } from '@assistant-ui/react';
import { ChatHeader } from '@/components/ChatHeader';
import { ChatPanel } from '@/components/ChatPanel';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';
import { useSettingsStore } from '@/stores';
import { useOfficeChat } from '@/hooks/useOfficeChat';
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

const ReadyAssistant: React.FC<{ host: OfficeHostApp }> = ({ host }) => {
  const { runtime, sessionError, isConnecting, clearMessages } = useOfficeChat(host);
  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <div className="flex h-screen flex-col overflow-hidden bg-background text-foreground">
        <ChatHeader onClearMessages={clearMessages} />
        {isConnecting && !sessionError && <ConnectingBanner />}
        {sessionError && <SessionErrorBanner error={sessionError} onRetry={clearMessages} />}
        <ChatErrorBoundary>
          <ChatPanel />
        </ChatErrorBoundary>
      </div>
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
