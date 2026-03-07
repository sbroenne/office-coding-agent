import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { CopilotModel, UserSettings } from '@/types';
import { DEFAULT_SETTINGS } from '@/types';
import { getAllAgents, setImportedAgents } from '@/services/agents';
import { setImportedSkills } from '@/services/skills';
import type { AgentSkill } from '@/types/skill';
import type { AgentConfig } from '@/types/agent';
import type { McpServerConfig } from '@/types/mcp';
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

  // ─── Skill management ───
  toggleSkill: (name: string) => void;
  isSkillEnabled: (name: string) => boolean;

  // ─── MCP server management ───
  toggleMcpServer: (name: string) => void;
  isMcpServerEnabled: (name: string) => boolean;

  // ─── Upload: imported skills ───
  addImportedSkill: (skill: AgentSkill) => void;
  removeImportedSkill: (name: string) => void;

  // ─── Upload: imported agents ───
  addImportedAgent: (agent: AgentConfig) => void;
  removeImportedAgent: (name: string) => void;

  // ─── Upload: imported MCP servers ───
  addImportedMcpServer: (server: McpServerConfig) => void;
  removeImportedMcpServer: (name: string) => void;

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
        // Sync agentService: plugin agents override same-named user-imported agents
        setImportedAgents([...agents, ...get().importedAgents]);
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

      // ─── Upload: imported skills ───
      addImportedSkill: skill => {
        const existing = get().importedSkills.filter(s => s.metadata.name !== skill.metadata.name);
        const updated = [...existing, skill];
        set({ importedSkills: updated });
        setImportedSkills(updated);
      },

      removeImportedSkill: name => {
        const updated = get().importedSkills.filter(s => s.metadata.name !== name);
        set({ importedSkills: updated });
        setImportedSkills(updated);
      },

      // ─── Upload: imported agents ───
      addImportedAgent: agent => {
        const existing = get().importedAgents.filter(a => a.metadata.name !== agent.metadata.name);
        const updated = [...existing, agent];
        set({ importedAgents: updated });
        setImportedAgents([...get().pluginAgents, ...updated]);
      },

      removeImportedAgent: name => {
        const updated = get().importedAgents.filter(a => a.metadata.name !== name);
        set({ importedAgents: updated });
        setImportedAgents([...get().pluginAgents, ...updated]);
      },

      // ─── Upload: imported MCP servers ───
      addImportedMcpServer: server => {
        const existing = get().importedMcpServers.filter(s => s.name !== server.name);
        set({ importedMcpServers: [...existing, server] });
      },

      removeImportedMcpServer: name => {
        set({ importedMcpServers: get().importedMcpServers.filter(s => s.name !== name) });
      },

      // ─── Reset ───
      reset: () => {
        set(DEFAULT_SETTINGS);
        setImportedSkills([]);
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
        importedSkills: state.importedSkills,
        importedAgents: state.importedAgents,
        importedMcpServers: state.importedMcpServers,
      }),
      onRehydrateStorage: () => state => {
        if (state) {
          setImportedSkills(state.importedSkills ?? []);
          setImportedAgents(state.importedAgents ?? []);
        }
      },
    }
  )
);
