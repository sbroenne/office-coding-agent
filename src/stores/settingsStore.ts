import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { CopilotModel, UserSettings } from '@/types';
import { DEFAULT_SETTINGS } from '@/types';
import { getAllAgents } from '@/services/agents';
import { officeStorage } from './officeStorage';

interface SettingsState extends UserSettings {
  // ─── Model management ───
  /** Models fetched from the Copilot SDK (cached across sessions) */
  availableModels: CopilotModel[] | null;
  setAvailableModels: (models: CopilotModel[]) => void;
  setActiveModel: (modelId: string) => void;

  // ─── Agent management ───
  setActiveAgent: (agentId: string) => void;
  getActiveAgent: () => string;

  // ─── Skill management ───
  toggleSkill: (name: string) => void;
  isSkillEnabled: (name: string) => boolean;

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

      // ─── Agent management ───
      setActiveAgent: agentId => {
        const agents = getAllAgents();
        const exists = agents.some(a => a.metadata.name === agentId);
        if (exists) {
          set({ activeAgentId: agentId });
        }
      },

      getActiveAgent: () => {
        return get().activeAgentId;
      },

      // ─── Skill management ───
      toggleSkill: name => {
        const disabled = get().disabledSkillNames;
        if (disabled.includes(name)) {
          set({ disabledSkillNames: disabled.filter(n => n !== name) });
        } else {
          set({ disabledSkillNames: [...disabled, name] });
        }
      },

      isSkillEnabled: name => {
        return !get().disabledSkillNames.includes(name);
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
        set({ ...DEFAULT_SETTINGS });
      },
    }),
    {
      name: 'office-coding-agent-settings',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        activeModel: state.activeModel,
        activeAgentId: state.activeAgentId,
        disabledSkillNames: state.disabledSkillNames,
        disabledMcpServerNames: state.disabledMcpServerNames,
      }),
    }
  )
);
