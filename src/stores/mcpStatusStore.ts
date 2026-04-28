import { create } from 'zustand';
import type { McpServerStatus, McpLogEntry, McpServerState } from '@/types';
import { toMcpServerKey } from '@/utils/mcpServerKey';

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

  /** Set transient OAuth UI state for a server */
  setOAuthState: (
    name: string,
    state: NonNullable<McpServerState['oauthState']>,
    alias?: string,
    error?: string
  ) => void;

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

function serverLookupKey(servers: Record<string, McpServerState>, name: string): string {
  const normalized = toMcpServerKey(name);
  return (
    Object.keys(servers).find(key => key === name || toMcpServerKey(key) === normalized) ??
    normalized
  );
}

export const useMcpStatusStore = create<McpStatusState>()(set => ({
  servers: {},

  setStatus: (name, status, error) => {
    set(state => {
      const key = serverLookupKey(state.servers, name);
      const servers = ensureServer(state.servers, key);
      const current = servers[key];
      const nextError =
        status === 'error' || status === 'failed' ? (error ?? current.error) : undefined;
      return {
        servers: {
          ...servers,
          [key]: {
            ...current,
            status,
            error: nextError,
            oauthState:
              status === 'connected'
                ? 'connected'
                : status === 'failed' || status === 'error'
                  ? 'failed'
                  : status === 'needs-auth'
                    ? 'idle'
                    : current.oauthState,
          },
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

  setOAuthState: (name, oauthState, alias, error) => {
    set(state => {
      const key = serverLookupKey(state.servers, name);
      const servers = ensureServer(state.servers, key);
      const current = servers[key];
      return {
        servers: {
          ...servers,
          [key]: {
            ...current,
            oauthState,
            oauthAlias: alias ?? current.oauthAlias,
            error: oauthState === 'failed' ? (error ?? current.error) : current.error,
          },
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
