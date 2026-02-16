import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import type { FoundryEndpoint, ModelInfo, UserSettings } from '@/types';
import { DEFAULT_SETTINGS } from '@/types';
import { generateId } from '@/utils/id';
import { invalidateClient } from '@/services/ai/aiClientFactory';
import { getSkills } from '@/services/skills';
import { getAllAgents } from '@/services/agents';
import { officeStorage } from './officeStorage';

interface SettingsState extends UserSettings {
  // ─── Endpoint management ───
  addEndpoint: (endpoint: Omit<FoundryEndpoint, 'id'>) => string;
  updateEndpoint: (id: string, updates: Partial<Omit<FoundryEndpoint, 'id'>>) => void;
  removeEndpoint: (id: string) => void;
  setActiveEndpoint: (id: string) => void;

  // ─── Model management (CRUD) ───
  addModel: (endpointId: string, model: ModelInfo) => void;
  updateModel: (
    endpointId: string,
    modelId: string,
    updates: Partial<Omit<ModelInfo, 'id'>>
  ) => void;
  removeModel: (endpointId: string, modelId: string) => void;
  setModelsForEndpoint: (endpointId: string, models: ModelInfo[]) => void;
  setActiveModel: (modelId: string) => void;

  // ─── Agent management ───
  setActiveAgent: (agentId: string) => void;
  getActiveAgent: () => string;

  // ─── Skill management ───
  toggleSkill: (skillName: string) => void;
  setActiveSkills: (skillNames: string[] | null) => void;
  getActiveSkillNames: () => string[] | null;

  // ─── Getters ───
  getActiveEndpoint: () => FoundryEndpoint | undefined;
  getActiveModel: () => ModelInfo | undefined;
  getModelsForActiveEndpoint: () => ModelInfo[];
  getModelsForEndpoint: (endpointId: string) => ModelInfo[];

  // ─── Reset ───
  reset: () => void;
}

export const useSettingsStore = create<SettingsState>()(
  persist(
    (set, get) => ({
      // ─── Initial state ───
      ...DEFAULT_SETTINGS,

      // ─── Endpoint management ───
      addEndpoint: endpoint => {
        const normalizedUrl = endpoint.resourceUrl.replace(/\/+$/, '');
        const existing = get().endpoints.find(
          ep => ep.resourceUrl.replace(/\/+$/, '') === normalizedUrl
        );

        if (existing) {
          // Update existing endpoint instead of creating a duplicate
          invalidateClient(existing.id);
          set(state => ({
            endpoints: state.endpoints.map(ep =>
              ep.id === existing.id ? { ...ep, ...endpoint, id: existing.id } : ep
            ),
            activeEndpointId: state.activeEndpointId ?? existing.id,
          }));
          return existing.id;
        }

        const id = generateId();
        const newEndpoint: FoundryEndpoint = { ...endpoint, id };
        set(state => ({
          endpoints: [...state.endpoints, newEndpoint],
          // Auto-activate if first endpoint
          activeEndpointId: state.activeEndpointId ?? id,
        }));
        return id;
      },

      updateEndpoint: (id, updates) => {
        invalidateClient(id); // Clear cached client so new config takes effect
        set(state => ({
          endpoints: state.endpoints.map(ep => (ep.id === id ? { ...ep, ...updates } : ep)),
        }));
      },

      removeEndpoint: id => {
        invalidateClient(id);
        set(state => {
          const remaining = state.endpoints.filter(ep => ep.id !== id);
          const { [id]: _removed, ...remainingModels } = state.endpointModels;
          return {
            endpoints: remaining,
            endpointModels: remainingModels,
            activeEndpointId:
              state.activeEndpointId === id ? (remaining[0]?.id ?? null) : state.activeEndpointId,
            activeModelId: state.activeEndpointId === id ? null : state.activeModelId,
          };
        });
      },

      setActiveEndpoint: id => {
        const models = get().endpointModels[id] ?? [];
        const defaultId = get().defaultModelId;
        const autoModel =
          models.find(m => m.id === defaultId) ?? models.find(m => m.isDefault) ?? models[0];
        set({
          activeEndpointId: id,
          activeModelId: autoModel?.id ?? null,
        });
      },

      // ─── Model management (CRUD) ───
      addModel: (endpointId, model) => {
        set(state => {
          const existing = state.endpointModels[endpointId] ?? [];
          // Avoid duplicates
          if (existing.some(m => m.id === model.id)) return state;
          const updated = [...existing, model];
          const patch: Partial<SettingsState> = {
            endpointModels: { ...state.endpointModels, [endpointId]: updated },
          };
          // Auto-select if this is the first model on the active endpoint
          if (state.activeEndpointId === endpointId && !state.activeModelId) {
            patch.activeModelId = model.id;
          }
          return patch;
        });
      },

      updateModel: (endpointId, modelId, updates) => {
        set(state => {
          const existing = state.endpointModels[endpointId] ?? [];
          return {
            endpointModels: {
              ...state.endpointModels,
              [endpointId]: existing.map(m => (m.id === modelId ? { ...m, ...updates } : m)),
            },
          };
        });
      },

      removeModel: (endpointId, modelId) => {
        set(state => {
          const remaining = (state.endpointModels[endpointId] ?? []).filter(m => m.id !== modelId);
          const patch: Partial<SettingsState> = {
            endpointModels: { ...state.endpointModels, [endpointId]: remaining },
          };
          // Clear active model if it was removed
          if (state.activeModelId === modelId && state.activeEndpointId === endpointId) {
            patch.activeModelId = remaining[0]?.id ?? null;
          }
          return patch;
        });
      },

      setModelsForEndpoint: (endpointId, models) => {
        set(state => {
          const patch: Partial<SettingsState> = {
            endpointModels: { ...state.endpointModels, [endpointId]: models },
          };
          // Auto-select default model if switching to this endpoint
          if (state.activeEndpointId === endpointId && !state.activeModelId) {
            const defaultModel =
              models.find(m => m.id === state.defaultModelId) ??
              models.find(m => m.isDefault) ??
              models[0];
            if (defaultModel) {
              patch.activeModelId = defaultModel.id;
            }
          }
          return patch;
        });
      },

      setActiveModel: modelId => {
        set({ activeModelId: modelId });
      },

      // ─── Agent management ───
      setActiveAgent: agentId => {
        // Validate agent exists
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
      toggleSkill: skillName => {
        set(state => {
          const current = state.activeSkillNames;
          if (current === null) {
            // All were on — materialize the full list minus the toggled one
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

      // ─── Getters ───
      getActiveEndpoint: () => {
        const state = get();
        return state.endpoints.find(ep => ep.id === state.activeEndpointId);
      },

      getActiveModel: () => {
        const state = get();
        if (!state.activeEndpointId || !state.activeModelId) return undefined;
        const models = state.endpointModels[state.activeEndpointId] ?? [];
        return models.find(m => m.id === state.activeModelId);
      },

      getModelsForActiveEndpoint: () => {
        const state = get();
        if (!state.activeEndpointId) return [];
        return state.endpointModels[state.activeEndpointId] ?? [];
      },

      getModelsForEndpoint: (endpointId: string) => {
        return get().endpointModels[endpointId] ?? [];
      },

      // ─── Reset ───
      reset: () => {
        set(DEFAULT_SETTINGS);
      },
    }),
    {
      name: 'office-coding-agent-settings',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        endpoints: state.endpoints,
        activeEndpointId: state.activeEndpointId,
        activeModelId: state.activeModelId,
        defaultModelId: state.defaultModelId,
        endpointModels: state.endpointModels,
        activeSkillNames: state.activeSkillNames,
        activeAgentId: state.activeAgentId,
      }),
    }
  )
);
