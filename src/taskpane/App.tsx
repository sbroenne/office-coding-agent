import React, { useEffect, useSyncExternalStore } from 'react';
import { Loader2 } from 'lucide-react';
import { AssistantRuntimeProvider } from '@assistant-ui/react';
import { ChatHeader } from '@/components/ChatHeader';
import { ChatPanel } from '@/components/ChatPanel';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';
import { useSettingsStore } from '@/stores';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { detectOfficeHost } from '@/services/office/host';
import type { OfficeHostApp } from '@/services/office/host';

const ReadyAssistant: React.FC<{ host: OfficeHostApp }> = ({ host }) => {
  const { runtime, clearMessages } = useOfficeChat(host);
  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <div className="flex h-screen flex-col overflow-hidden bg-background text-foreground">
        <ChatHeader onClearMessages={clearMessages} />
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
