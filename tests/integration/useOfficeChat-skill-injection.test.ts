/**
 * Integration test: useOfficeChat — imported skill injection.
 *
 * Regression test for the `importedSkills` memo-dep bug:
 * When a skill is imported while `activeSkillNames` is null (all ON),
 * the agent must rebuild with the new skill's content in its instructions.
 *
 * Root cause: `useMemo` for `agent` in useOfficeChat did not include
 * `importedSkills` in its dep array, so importing a skill (which doesn't
 * change `activeSkillNames`) never triggered a rebuild.
 *
 * Fix: Added `importedSkills` to the dep array.
 */

import { renderHook, act } from '@testing-library/react';
import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { LanguageModel } from 'ai';
import type { AgentSkill } from '@/types';
import { useOfficeChat } from '@/hooks/useOfficeChat';
import { useSettingsStore } from '@/stores/settingsStore';
import { setImportedSkills } from '@/services/skills/skillService';

// ─── Mocks ─────────────────────────────────────────────────────────────────────

// Capture instructions each time ToolLoopAgent is constructed.
const constructedInstructions: string[] = [];

vi.mock('ai', async importOriginal => {
  const real = (await importOriginal()) as Record<string, unknown>;
  return {
    ...real,
    ToolLoopAgent: class MockToolLoopAgent {
      constructor(opts: { instructions: string }) {
        constructedInstructions.push(opts.instructions);
      }
    },
    DirectChatTransport: class MockDirectChatTransport {
      constructor() {}
    },
  };
});

vi.mock('@ai-sdk/react', () => ({
  useChat: vi.fn(() => ({})),
}));

vi.mock('@/hooks/useMcpTools', () => ({
  useMcpTools: vi.fn(() => ({})),
}));

// ─── Fixtures ──────────────────────────────────────────────────────────────────

const fakeModel = {} as unknown as LanguageModel;

const importedSkill: AgentSkill = {
  metadata: {
    name: 'My Custom Skill',
    description: 'A custom skill for testing.',
    version: '1.0.0',
    tags: [],
  },
  content: 'Custom skill content injected here.',
};

// ─── Tests ─────────────────────────────────────────────────────────────────────

describe('useOfficeChat — imported skill injection', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    constructedInstructions.length = 0;
    useSettingsStore.getState().reset();
    setImportedSkills([]);
  });

  it('does not include imported skill content before the skill is imported', () => {
    renderHook(() => useOfficeChat(fakeModel, 'excel'));

    expect(constructedInstructions.at(-1)).not.toContain('My Custom Skill');
    expect(constructedInstructions.at(-1)).not.toContain('Custom skill content injected here.');
  });

  it('rebuilds agent instructions when a new skill is imported (null activeSkillNames = all ON)', () => {
    const { rerender } = renderHook(() => useOfficeChat(fakeModel, 'excel'));

    // Confirm activeSkillNames is null (all ON) — importing should be immediately active
    expect(useSettingsStore.getState().activeSkillNames).toBeNull();

    const callsBefore = constructedInstructions.length;
    expect(constructedInstructions.at(-1)).not.toContain('My Custom Skill');

    // Import the skill via the store action (simulates user importing a .md file)
    act(() => {
      useSettingsStore.getState().importSkills([importedSkill]);
    });
    rerender();

    // Agent must have been rebuilt (new constructor call)
    expect(constructedInstructions.length).toBeGreaterThan(callsBefore);
    // New instructions must include the imported skill's content
    expect(constructedInstructions.at(-1)).toContain('My Custom Skill');
    expect(constructedInstructions.at(-1)).toContain('Custom skill content injected here.');
  });

  it('removing the skill causes agent to rebuild without its content', () => {
    const { rerender } = renderHook(() => useOfficeChat(fakeModel, 'excel'));

    act(() => {
      useSettingsStore.getState().importSkills([importedSkill]);
    });
    rerender();
    expect(constructedInstructions.at(-1)).toContain('My Custom Skill');

    act(() => {
      useSettingsStore.getState().removeImportedSkill('My Custom Skill');
    });
    rerender();

    expect(constructedInstructions.at(-1)).not.toContain('My Custom Skill');
    expect(constructedInstructions.at(-1)).not.toContain('Custom skill content injected here.');
  });
});
