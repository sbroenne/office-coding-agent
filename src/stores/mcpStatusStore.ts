import { create } from 'zustand';
import type { McpServerStatus, McpLogEntry, McpServerState } from '@/types';

const MAX_LOGS_PER_SERVER = 500;

interface McpStatusState {
  /** Per-server runtime state keyed by server name */
  servers: Record<string, McpServerState>;

  /** Set the status of a server */
  setStatus: (name: string, status: McpServerStatus, error?: string) => void;

  /** Append a log entry for a server */
  addLog: (name: string, entry: McpLogEntry) => void;

  /** Set the discovered tools for a server */
  setTools: (name: string, tools: { name: string; description: string }[]) => void;

  /** Clear logs for a specific server */
  clearLogs: (name: string) => void;

  /** Reset all server state (e.g. on session reinit) */
  clearAll: () => void;
}

function defaultServerState(): McpServerState {
  return { status: 'stopped', tools: [], logs: [] };
}

function ensureServer(
  servers: Record<string, McpServerState>,
  name: string
): Record<string, McpServerState> {
  if (servers[name]) return servers;
  return { ...servers, [name]: defaultServerState() };
}

export const useMcpStatusStore = create<McpStatusState>()(set => ({
  servers: {},

  setStatus: (name, status, error) => {
    set(state => {
      const servers = ensureServer(state.servers, name);
      const current = servers[name];
      return {
        servers: {
          ...servers,
          [name]: { ...current, status, error: error ?? undefined },
        },
      };
    });
  },

  addLog: (name, entry) => {
    set(state => {
      const servers = ensureServer(state.servers, name);
      const current = servers[name];
      const logs = [...current.logs, entry];
      // Trim to max
      const trimmed =
        logs.length > MAX_LOGS_PER_SERVER ? logs.slice(logs.length - MAX_LOGS_PER_SERVER) : logs;
      return {
        servers: {
          ...servers,
          [name]: { ...current, logs: trimmed },
        },
      };
    });
  },

  setTools: (name, tools) => {
    set(state => {
      const servers = ensureServer(state.servers, name);
      const current = servers[name];
      return {
        servers: {
          ...servers,
          [name]: { ...current, tools },
        },
      };
    });
  },

  clearLogs: name => {
    set(state => {
      const current = state.servers[name];
      if (!current) return state;
      return {
        servers: {
          ...state.servers,
          [name]: { ...current, logs: [] },
        },
      };
    });
  },

  clearAll: () => {
    set({ servers: {} });
  },
}));
