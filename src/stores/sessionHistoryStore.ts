import { create } from 'zustand';
import { createJSONStorage, persist } from 'zustand/middleware';
import { officeStorage } from './officeStorage';
import { generateId } from '@/utils/id';
import type { OfficeHostApp } from '@/services/office/host';

export interface SessionHistoryItem {
  id: string;
  title: string;
  host: OfficeHostApp;
  updatedAt: number;
  messages: unknown[];
}

interface SessionHistoryStoreState {
  sessions: SessionHistoryItem[];
  activeSessionId: string | null;
  createSession: (host: OfficeHostApp) => string;
  setActiveSession: (sessionId: string) => void;
  upsertActiveSession: (input: { host: OfficeHostApp; title: string; messages: unknown[] }) => void;
  deleteSession: (sessionId: string) => void;
  clearSessionsForHost: (host: OfficeHostApp) => void;
}

const MAX_SESSIONS = 50;

function toSerializableMessages(messages: unknown[]): unknown[] {
  try {
    return JSON.parse(JSON.stringify(messages)) as unknown[];
  } catch {
    return [];
  }
}

function trimSessions(items: SessionHistoryItem[]): SessionHistoryItem[] {
  return [...items].sort((a, b) => b.updatedAt - a.updatedAt).slice(0, MAX_SESSIONS);
}

export const useSessionHistoryStore = create<SessionHistoryStoreState>()(
  persist(
    (set, get) => ({
      sessions: [],
      activeSessionId: null,

      createSession: host => {
        const id = generateId();
        const now = Date.now();
        const item: SessionHistoryItem = {
          id,
          title: 'New conversation',
          host,
          updatedAt: now,
          messages: [],
        };

        set(state => ({
          sessions: trimSessions([item, ...state.sessions]),
          activeSessionId: id,
        }));

        return id;
      },

      setActiveSession: sessionId => {
        set({ activeSessionId: sessionId });
      },

      upsertActiveSession: ({ host, title, messages }) => {
        const state = get();
        const sessionId = state.activeSessionId ?? state.createSession(host);
        const now = Date.now();
        const next: SessionHistoryItem = {
          id: sessionId,
          host,
          title,
          updatedAt: now,
          messages: toSerializableMessages(messages),
        };

        set(current => {
          const remaining = current.sessions.filter(s => s.id !== sessionId);
          return {
            sessions: trimSessions([next, ...remaining]),
            activeSessionId: sessionId,
          };
        });
      },

      deleteSession: sessionId => {
        set(state => {
          const nextSessions = state.sessions.filter(s => s.id !== sessionId);
          const nextActiveSessionId =
            state.activeSessionId === sessionId
              ? (nextSessions[0]?.id ?? null)
              : state.activeSessionId;
          return {
            sessions: nextSessions,
            activeSessionId: nextActiveSessionId,
          };
        });
      },

      clearSessionsForHost: host => {
        set(state => {
          const nextSessions = state.sessions.filter(s => s.host !== host);
          const activeStillExists = nextSessions.some(s => s.id === state.activeSessionId);
          return {
            sessions: nextSessions,
            activeSessionId: activeStillExists
              ? state.activeSessionId
              : (nextSessions[0]?.id ?? null),
          };
        });
      },
    }),
    {
      name: 'office-coding-agent-session-history',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        sessions: state.sessions,
        activeSessionId: state.activeSessionId,
      }),
    }
  )
);
