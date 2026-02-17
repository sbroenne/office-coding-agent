import React, { useEffect, useState, useCallback, useSyncExternalStore, useMemo, useRef } from 'react';
import { Loader2 } from 'lucide-react';
import { AssistantRuntimeProvider } from '@assistant-ui/react';
import { useAISDKRuntime } from '@assistant-ui/react-ai-sdk';
import type { AzureOpenAIProvider } from '@ai-sdk/azure';
import { ChatHeader } from '@/components/ChatHeader';
import { ChatPanel } from '@/components/ChatPanel';
import { ChatErrorBoundary } from '@/components/ChatErrorBoundary';
import { SetupWizard } from '@/components/SetupWizard';
import { useChatStore, useSettingsStore } from '@/stores';
import { getAzureProvider } from '@/services/ai/aiClientFactory';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { detectOfficeHost, type OfficeHostApp } from '@/services/office/host';

type AppState = 'loading' | 'setup' | 'ready';

interface ReadyAssistantProps {
  provider: AzureOpenAIProvider;
  modelId: string;
  host: OfficeHostApp;
  settingsOpen: boolean;
  onSettingsOpenChange: (open: boolean) => void;
  onOpenSettings: () => void;
}

const ReadyAssistant: React.FC<ReadyAssistantProps> = ({
  provider,
  modelId,
  host,
  settingsOpen,
  onSettingsOpenChange,
  onOpenSettings,
}) => {
  const chat = useOfficeChat(provider, modelId, host);
  const persistedMessages = useChatStore(s => s.messages);
  const setPersistedMessages = useChatStore(s => s.setMessages);
  const clearPersistedMessages = useChatStore(s => s.clearMessages);
  const hydratedOnceRef = useRef(false);
  const hasHydratedChat = useSyncExternalStore(useChatStore.persist.onFinishHydration, () =>
    useChatStore.persist.hasHydrated()
  );
  const runtime = useAISDKRuntime(chat);

  useEffect(() => {
    if (!hasHydratedChat || hydratedOnceRef.current) return;
    hydratedOnceRef.current = true;
    chat.setMessages((persistedMessages as typeof chat.messages) ?? []);
  }, [chat, hasHydratedChat, persistedMessages]);

  useEffect(() => {
    if (!hasHydratedChat) return;
    const currentSerialized = JSON.stringify(chat.messages ?? []);
    const persistedSerialized = JSON.stringify(persistedMessages ?? []);
    if (currentSerialized === persistedSerialized) return;
    setPersistedMessages(chat.messages as unknown[]);
  }, [chat.messages, hasHydratedChat, persistedMessages, setPersistedMessages]);

  const handleClearMessages = useCallback(() => {
    chat.setMessages([]);
    clearPersistedMessages();
  }, [chat, clearPersistedMessages]);

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <div className="flex h-screen flex-col overflow-hidden bg-background text-foreground">
        <ChatHeader
          onClearMessages={handleClearMessages}
          settingsOpen={settingsOpen}
          onSettingsOpenChange={onSettingsOpenChange}
        />
        <ChatErrorBoundary>
          <ChatPanel isConfigured={true} onOpenSettings={onOpenSettings} />
        </ChatErrorBoundary>
      </div>
    </AssistantRuntimeProvider>
  );
};

export const App: React.FC = () => {
  const [appState, setAppState] = useState<AppState>('loading');
  const endpoints = useSettingsStore(s => s.endpoints);
  const activeEndpointId = useSettingsStore(s => s.activeEndpointId);
  const activeModelId = useSettingsStore(s => s.activeModelId);
  const endpointModels = useSettingsStore(s => s.endpointModels);

  // Wait for Zustand persist to finish hydrating from async storage.
  // Without this, the store starts with DEFAULT_SETTINGS (empty endpoints)
  // and the app immediately shows SetupWizard before saved data loads.
  const hasHydrated = useSyncExternalStore(useSettingsStore.persist.onFinishHydration, () =>
    useSettingsStore.persist.hasHydrated()
  );

  // Detect system theme preference, reacting to OS changes
  const prefersDark = useSyncExternalStore(
    onStoreChange => {
      if (typeof window === 'undefined') {
        // SSR — no matchMedia available
        return () => undefined;
      }
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

  const activeModels = activeEndpointId ? (endpointModels[activeEndpointId] ?? []) : [];
  const needsSetup = endpoints.length === 0 || activeModels.length === 0;

  // Create the Azure provider from the active endpoint (memoised).
  // Depends on endpoints array + activeEndpointId; getAzureProvider caches internally.
  const provider = useMemo(() => {
    const ep = endpoints.find(e => e.id === activeEndpointId);
    return ep ? getAzureProvider(ep) : null;
  }, [endpoints, activeEndpointId]);

  const host = useMemo(() => detectOfficeHost(), []);

  const [settingsOpen, setSettingsOpen] = useState(false);
  const handleOpenSettings = useCallback(() => setSettingsOpen(true), []);

  // Once hydrated, transition from 'loading' to the correct initial state.
  // Include needsSetup in deps so we re-evaluate if the store fires its
  // subscribers slightly after the hydration callback.
  useEffect(() => {
    if (!hasHydrated || appState !== 'loading') return;
    setAppState(needsSetup ? 'setup' : 'ready');
  }, [hasHydrated, needsSetup, appState]);

  // React to store changes: if config is wiped while chatting, go back to setup
  useEffect(() => {
    if (appState !== 'ready') return;
    if (needsSetup) {
      setAppState('setup');
    }
  }, [appState, needsSetup]);

  const handleSetupComplete = useCallback(() => {
    setAppState('ready');
  }, []);

  if (appState === 'loading') {
    return (
      <div className="flex h-screen flex-col items-center justify-center gap-3 bg-background text-foreground">
        <Loader2 className="size-8 animate-spin text-muted-foreground" />
        <p className="text-sm text-muted-foreground">Initializing...</p>
      </div>
    );
  }

  if (appState === 'setup') {
    return <SetupWizard onComplete={handleSetupComplete} />;
  }

  if (!provider || !activeModelId) {
    return <SetupWizard onComplete={handleSetupComplete} />;
  }

  return (
    <ReadyAssistant
      provider={provider}
      modelId={activeModelId}
      host={host}
      settingsOpen={settingsOpen}
      onSettingsOpenChange={setSettingsOpen}
      onOpenSettings={handleOpenSettings}
    />
  );
};
