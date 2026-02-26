import { create } from 'zustand';
import { createJSONStorage, persist } from 'zustand/middleware';
import { officeStorage } from './officeStorage';
import type { PermissionRequestPayload } from '@/lib/websocket-client';

export interface PermissionRule {
  id: string;
  kind: string;
  pathPrefix: string;
}

interface PermissionStoreState {
  allowAll: boolean;
  workingDirectory: string | null;
  rules: PermissionRule[];
  setAllowAll: (value: boolean) => void;
  setWorkingDirectory: (value: string | null) => void;
  addRule: (rule: Omit<PermissionRule, 'id'>) => void;
  removeRule: (id: string) => void;
  clearRules: () => void;
  evaluate: (request: PermissionRequestPayload['request']) => 'approved' | null;
}

function normalizePath(value: string): string {
  return value.replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
}

function isUnderPath(candidate: string, prefix: string): boolean {
  const normalizedCandidate = normalizePath(candidate);
  const normalizedPrefix = normalizePath(prefix);
  return (
    normalizedCandidate === normalizedPrefix ||
    normalizedCandidate.startsWith(`${normalizedPrefix}/`)
  );
}

function pathForRequest(request: PermissionRequestPayload['request']): string | null {
  if (typeof request.path === 'string' && request.path.length > 0) return request.path;
  if (typeof request.fileName === 'string' && request.fileName.length > 0) return request.fileName;
  if (typeof request.fullCommandText === 'string' && request.fullCommandText.length > 0)
    return request.fullCommandText;
  return null;
}

export const usePermissionStore = create<PermissionStoreState>()(
  persist(
    (set, get) => ({
      allowAll: true,
      workingDirectory: null,
      rules: [],

      setAllowAll: value => {
        set({ allowAll: value });
      },

      setWorkingDirectory: value => {
        set({ workingDirectory: value });
      },

      addRule: rule => {
        const id = `${rule.kind}:${normalizePath(rule.pathPrefix)}`;
        set(state => {
          if (state.rules.some(r => r.id === id)) return state;
          return {
            rules: [...state.rules, { id, kind: rule.kind, pathPrefix: rule.pathPrefix }],
          };
        });
      },

      removeRule: id => {
        set(state => ({ rules: state.rules.filter(rule => rule.id !== id) }));
      },

      clearRules: () => {
        set({ rules: [] });
      },

      evaluate: request => {
        const state = get();

        if (state.allowAll) return 'approved';

        const requestPath = pathForRequest(request);

        if (
          request.kind === 'read' &&
          requestPath &&
          state.workingDirectory &&
          isUnderPath(requestPath, state.workingDirectory)
        ) {
          return 'approved';
        }

        for (const rule of state.rules) {
          if (rule.kind !== request.kind) continue;
          if (!requestPath) continue;
          if (isUnderPath(requestPath, rule.pathPrefix)) return 'approved';
        }

        return null;
      },
    }),
    {
      name: 'office-coding-agent-permissions',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        allowAll: state.allowAll,
        workingDirectory: state.workingDirectory,
        rules: state.rules,
      }),
    }
  )
);
