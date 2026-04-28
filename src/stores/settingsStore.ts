import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { CopilotAgent, CopilotModel, UserSettings } from '@/types';
import { DEFAULT_SETTINGS } from '@/types';
import { officeStorage } from './officeStorage';

interface SettingsState extends UserSettings {
  // ─── Model management ───
  /** Models fetched from the Copilot SDK (cached across sessions) */
  availableModels: CopilotModel[] | null;
  setAvailableModels: (models: CopilotModel[]) => void;
  setActiveModel: (modelId: string) => void;

  // ─── CLI agent management ───
  /** Agents discovered from the Copilot CLI (cached across sessions) */
  availableAgents: CopilotAgent[] | null;
  setAvailableAgents: (agents: CopilotAgent[]) => void;
  setActiveAgent: (agentName: string | null) => void;

  // ─── MCP server management ───
  toggleMcpServer: (name: string) => void;
  isMcpServerEnabled: (name: string) => boolean;

  // ─── Reset ───
  reset: () => void;
}

export const useSettingsStore = create<SettingsState>()(
  persist(
    (set, get) => ({
      // ─── Initial state ───
      ...DEFAULT_SETTINGS,
      availableModels: null,
      availableAgents: null,

      // ─── Model management ───
      setAvailableModels: models => {
        set({ availableModels: models });
      },

      setActiveModel: modelId => {
        const models = get().availableModels;
        if (!models || models.some(m => m.id === modelId)) {
          set({ activeModel: modelId });
        }
      },

      // ─── CLI agent management ───
      setAvailableAgents: agents => {
        set({ availableAgents: agents });
      },

      setActiveAgent: agentName => {
        const agents = get().availableAgents;
        if (agentName === null || !agents || agents.some(a => a.name === agentName)) {
          set({ activeAgentName: agentName });
        }
      },

      // ─── MCP server management ───
      toggleMcpServer: name => {
        const disabled = get().disabledMcpServerNames;
        if (disabled.includes(name)) {
          set({ disabledMcpServerNames: disabled.filter(n => n !== name) });
        } else {
          set({ disabledMcpServerNames: [...disabled, name] });
        }
      },

      isMcpServerEnabled: name => {
        return !get().disabledMcpServerNames.includes(name);
      },

      // ─── Reset ───
      reset: () => {
        set({ ...DEFAULT_SETTINGS, availableModels: null, availableAgents: null });
      },
    }),
    {
      name: 'office-coding-agent-settings',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        activeModel: state.activeModel,
        activeAgentName: state.activeAgentName,
        disabledMcpServerNames: state.disabledMcpServerNames,
      }),
    }
  )
);
