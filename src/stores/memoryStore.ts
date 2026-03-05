import { create } from 'zustand';
import { persist, createJSONStorage } from 'zustand/middleware';
import { officeStorage } from './officeStorage';

/** A single memory entry stored by the agent */
export interface MemoryEntry {
  /** Unique ID */
  id: string;
  /** The fact or preference to remember */
  content: string;
  /** Optional category for organization */
  category?: string;
  /** When this memory was created */
  createdAt: number;
  /** When this memory was last accessed/used */
  lastUsedAt: number;
}

interface MemoryState {
  memories: MemoryEntry[];

  /** Add or update a memory. If content matches an existing one, updates it. */
  addMemory: (content: string, category?: string) => string;

  /** Remove a memory by ID */
  removeMemory: (id: string) => void;

  /** List all memories, optionally filtered by category */
  listMemories: (category?: string) => MemoryEntry[];

  /** Search memories by keyword */
  searchMemories: (query: string) => MemoryEntry[];

  /** Clear all memories */
  clearMemories: () => void;

  /** Build a context string for injection into the system prompt */
  buildMemoryContext: () => string;
}

const MAX_MEMORIES = 100;

export const useMemoryStore = create<MemoryState>()(
  persist(
    (set, get) => ({
      memories: [],

      addMemory: (content: string, category?: string) => {
        const trimmed = content.trim();
        if (!trimmed) return '';

        const existing = get().memories.find(
          m => m.content.toLowerCase() === trimmed.toLowerCase()
        );

        if (existing) {
          set(state => ({
            memories: state.memories.map(m =>
              m.id === existing.id ? { ...m, lastUsedAt: Date.now(), category: category ?? m.category } : m
            ),
          }));
          return existing.id;
        }

        const id = `mem-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        const entry: MemoryEntry = {
          id,
          content: trimmed,
          category,
          createdAt: Date.now(),
          lastUsedAt: Date.now(),
        };

        set(state => {
          const updated = [entry, ...state.memories];
          // Keep only the most recent MAX_MEMORIES
          if (updated.length > MAX_MEMORIES) {
            updated.sort((a, b) => b.lastUsedAt - a.lastUsedAt);
            updated.length = MAX_MEMORIES;
          }
          return { memories: updated };
        });

        return id;
      },

      removeMemory: (id: string) => {
        set(state => ({
          memories: state.memories.filter(m => m.id !== id),
        }));
      },

      listMemories: (category?: string) => {
        const all = get().memories;
        if (!category) return all;
        return all.filter(m => m.category?.toLowerCase() === category.toLowerCase());
      },

      searchMemories: (query: string) => {
        const lower = query.toLowerCase();
        return get().memories.filter(m =>
          m.content.toLowerCase().includes(lower) ||
          (m.category?.toLowerCase().includes(lower) ?? false)
        );
      },

      clearMemories: () => set({ memories: [] }),

      buildMemoryContext: () => {
        const memories = get().memories;
        if (memories.length === 0) return '';

        const grouped = new Map<string, string[]>();
        for (const m of memories) {
          const cat = m.category ?? 'general';
          const list = grouped.get(cat) ?? [];
          list.push(m.content);
          grouped.set(cat, list);
        }

        const sections: string[] = [];
        for (const [cat, items] of grouped) {
          sections.push(`**${cat}**:\n${items.map(i => `- ${i}`).join('\n')}`);
        }

        return `## User Memories\n\nThese are facts and preferences you've learned about this user. Use them to personalize your responses.\n\n${sections.join('\n\n')}`;
      },
    }),
    {
      name: 'office-coding-agent-memories',
      storage: createJSONStorage(() => officeStorage),
    }
  )
);
