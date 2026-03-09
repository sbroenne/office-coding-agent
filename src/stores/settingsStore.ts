import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { CopilotModel, UserSettings } from '@/types';
import { DEFAULT_SETTINGS } from '@/types';
import { getAllAgents, setImportedAgents } from '@/services/agents';
import type { AgentSkill } from '@/types/skill';
import type { AgentConfig } from '@/types/agent';
import type { PluginPrompt } from '@/types/plugin';
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
  /**
   * Plugin agents discovered from the Copilot CLI config at session start.
   * Ephemeral — NOT persisted. Populated by the plugin.agents notification.
   */
  pluginAgents: AgentConfig[];
  /** Replace the current plugin agent list (called on each session.create). */
  setPluginAgents: (agents: AgentConfig[]) => void;

  /**
   * Plugin skills discovered from the Copilot CLI config at session start.
   * Ephemeral — NOT persisted. Populated by the plugin.skills notification.
   */
  pluginSkills: AgentSkill[];
  /** Replace the current plugin skill list (called on each session.create). */
  setPluginSkills: (skills: AgentSkill[]) => void;

  /**
   * Plugin prompts (slash commands) discovered from plugin prompts/ directories.
   * Ephemeral — NOT persisted. Populated by the plugin.prompts notification.
   */
  pluginPrompts: PluginPrompt[];
  /** Replace the current plugin prompt list (called on each session.create). */
  setPluginPrompts: (prompts: PluginPrompt[]) => void;

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
      pluginAgents: [],
      pluginSkills: [],
      pluginPrompts: [],

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

      setPluginAgents: agents => {
        set({ pluginAgents: agents });
        setImportedAgents(agents);
      },

      setPluginSkills: skills => {
        set({ pluginSkills: skills });
      },

      setPluginPrompts: prompts => {
        set({ pluginPrompts: prompts });
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
        set({ ...DEFAULT_SETTINGS, pluginAgents: [], pluginSkills: [], pluginPrompts: [] });
        setImportedAgents([]);
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
