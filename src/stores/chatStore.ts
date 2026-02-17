import { create } from 'zustand';
import { createJSONStorage, persist } from 'zustand/middleware';
import { officeStorage } from './officeStorage';

interface ChatStoreState {
  messages: unknown[];
  setMessages: (messages: unknown[]) => void;
  clearMessages: () => void;
  reset: () => void;
}

const DEFAULT_CHAT_STATE = {
  messages: [],
} as const;

function toSerializableMessages(messages: unknown[]): unknown[] {
  try {
    return JSON.parse(JSON.stringify(messages)) as unknown[];
  } catch {
    return [];
  }
}

export const useChatStore = create<ChatStoreState>()(
  persist(
    set => ({
      ...DEFAULT_CHAT_STATE,

      setMessages: messages => {
        set({ messages: toSerializableMessages(messages) });
      },

      clearMessages: () => {
        set({ messages: [] });
      },

      reset: () => {
        set({ ...DEFAULT_CHAT_STATE });
      },
    }),
    {
      name: 'office-coding-agent-chat',
      storage: createJSONStorage(() => officeStorage),
      partialize: state => ({
        messages: state.messages,
      }),
    }
  )
);
