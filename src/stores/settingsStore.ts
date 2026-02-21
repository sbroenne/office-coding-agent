import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { CopilotModel, UserSettings } from '@/types';
import { DEFAULT_SETTINGS } from '@/types';
import { getAllAgents, getBundledAgents, setImportedAgents } from '@/services/agents';
import { getBundledSkills, getSkills, setImportedSkills } from '@/services/skills';
import type { AgentConfig, AgentSkill, McpServerConfig } from '@/types';
import { officeStorage } from './officeStorage';

function ensureUniqueImportedName(baseName: string, existingNames: Set<string>): string {
  if (!existingNames.has(baseName)) return baseName;

  let index = 1;
  let candidate = `${baseName} (imported)`;
  while (existingNames.has(candidate)) {
    index += 1;
    candidate = `${baseName} (imported ${index})`;
  }

  return candidate;
}

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
  toggleSkill: (skillName: string) => void;
  setActiveSkills: (skillNames: string[] | null) => void;
  getActiveSkillNames: () => string[] | null;
  importSkills: (skills: AgentSkill[]) => void;
  removeImportedSkill: (skillName: string) => void;

  // ─── Imported agent management ───
  importAgents: (agents: AgentConfig[]) => void;
  removeImportedAgent: (agentName: string) => void;

  // ─── MCP server management ───
  importMcpServers: (servers: McpServerConfig[]) => void;
  removeMcpServer: (serverName: string) => void;
  toggleMcpServer: (serverName: string) => void;

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

      importAgents: agents => {
        set(state => {
          const existingNames = new Set([
            ...getBundledAgents().map(agent => agent.metadata.name),
            ...state.importedAgents.map(agent => agent.metadata.name),
          ]);

          const nextImported = [...state.importedAgents];
          for (const agent of agents) {
            const uniqueName = ensureUniqueImportedName(agent.metadata.name, existingNames);
            existingNames.add(uniqueName);
            nextImported.push({
              ...agent,
              metadata: {
                ...agent.metadata,
                name: uniqueName,
              },
            });
          }

          setImportedAgents(nextImported);
          return { importedAgents: nextImported };
        });
      },

      removeImportedAgent: agentName => {
        set(state => {
          const nextImported = state.importedAgents.filter(
            agent => agent.metadata.name !== agentName
          );
          setImportedAgents(nextImported);

          const nextActiveAgentId =
            state.activeAgentId === agentName
              ? DEFAULT_SETTINGS.activeAgentId
              : state.activeAgentId;

          return {
            importedAgents: nextImported,
            activeAgentId: nextActiveAgentId,
          };
        });
      },

      // ─── Skill management ───
      toggleSkill: skillName => {
        set(state => {
          const current = state.activeSkillNames;
          if (current === null) {
            const allNames = getSkills().map(s => s.metadata.name);
            return { activeSkillNames: allNames.filter(n => n !== skillName) };
          }
          const next = current.includes(skillName)
            ? current.filter(n => n !== skillName)
            : [...current, skillName];
          return { activeSkillNames: next };
        });
      },

      setActiveSkills: skillNames => {
        set({ activeSkillNames: skillNames });
      },

      getActiveSkillNames: () => {
        return get().activeSkillNames;
      },

      importSkills: skills => {
        set(state => {
          const existingNames = new Set([
            ...getBundledSkills().map(skill => skill.metadata.name),
            ...state.importedSkills.map(skill => skill.metadata.name),
          ]);

          const nextImported = [...state.importedSkills];
          for (const skill of skills) {
            const uniqueName = ensureUniqueImportedName(skill.metadata.name, existingNames);
            existingNames.add(uniqueName);
            nextImported.push({
              ...skill,
              metadata: {
                ...skill.metadata,
                name: uniqueName,
              },
            });
          }

          setImportedSkills(nextImported);
          return { importedSkills: nextImported };
        });
      },

      removeImportedSkill: skillName => {
        set(state => {
          const nextImported = state.importedSkills.filter(
            skill => skill.metadata.name !== skillName
          );
          setImportedSkills(nextImported);

          const nextActiveSkillNames =
            state.activeSkillNames?.filter(name => name !== skillName) ?? null;

          return {
            importedSkills: nextImported,
            activeSkillNames: nextActiveSkillNames,
          };
        });
      },

      // ─── MCP server management ───
      importMcpServers: servers => {
        set(state => {
          const existingNames = new Set(state.importedMcpServers.map(s => s.name));
          const nextImported = [...state.importedMcpServers];
          for (const server of servers) {
            const uniqueName = ensureUniqueImportedName(server.name, existingNames);
            existingNames.add(uniqueName);
            nextImported.push({ ...server, name: uniqueName });
          }
          return { importedMcpServers: nextImported };
        });
      },

      removeMcpServer: serverName => {
        set(state => {
          const nextImported = state.importedMcpServers.filter(s => s.name !== serverName);
          const nextActiveNames = state.activeMcpServerNames?.filter(n => n !== serverName) ?? null;
          return { importedMcpServers: nextImported, activeMcpServerNames: nextActiveNames };
        });
      },

      toggleMcpServer: serverName => {
        set(state => {
          const current = state.activeMcpServerNames;
          if (current === null) {
            const allNames = state.importedMcpServers.map(s => s.name);
            return { activeMcpServerNames: allNames.filter(n => n !== serverName) };
          }
          const next = current.includes(serverName)
            ? current.filter(n => n !== serverName)
            : [...current, serverName];
          return { activeMcpServerNames: next };
        });
      },

      // ─── Reset ───
      reset: () => {
        setImportedSkills([]);
        setImportedAgents([]);
        set(DEFAULT_SETTINGS);
      },
    }),
    {
      name: 'office-coding-agent-settings',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        activeModel: state.activeModel,
        activeSkillNames: state.activeSkillNames,
        activeAgentId: state.activeAgentId,
        importedSkills: state.importedSkills,
        importedAgents: state.importedAgents,
        importedMcpServers: state.importedMcpServers,
        activeMcpServerNames: state.activeMcpServerNames,
        // availableModels is NOT persisted — it's always fetched fresh from the
        // Copilot CLI on connect, so a stale cached list never survives restarts.
      }),
      onRehydrateStorage: () => state => {
        setImportedSkills(state?.importedSkills ?? []);
        setImportedAgents(state?.importedAgents ?? []);
      },
    }
  )
);
