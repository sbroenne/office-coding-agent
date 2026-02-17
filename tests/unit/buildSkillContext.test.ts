/**
 * Unit tests for buildSkillContext and related skill functions.
 *
 * These exercise the real `skillService` module which imports
 * bundled `.md` skill files via the rawMarkdownPlugin in vitest.config.ts.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import type { AgentSkill } from '@/types/skill';
import {
  buildSkillContext,
  getSkills,
  getSkill,
  parseFrontmatter,
  setImportedSkills,
} from '@/services/skills/skillService';

beforeEach(() => {
  setImportedSkills([]);
});

afterEach(() => {
  setImportedSkills([]);
});

describe('buildSkillContext', () => {
  it('includes bundled skills when available', () => {
    const ctx = buildSkillContext();
    expect(ctx).toContain('# Agent Skills');
    expect(ctx).toContain('Agent Skill: excel');
  });

  it('returns empty string when activeNames do not match any skills', () => {
    const ctx = buildSkillContext(['any-skill-that-does-not-exist']);
    expect(ctx).toBe('');
  });

  it('returns empty string when activeNames is an empty array', () => {
    const ctx = buildSkillContext([]);
    expect(ctx).toBe('');
  });

  it('returns empty string when no names match', () => {
    const ctx = buildSkillContext(['nonexistent-skill']);
    expect(ctx).toBe('');
  });

  it('includes all skills when activeNames is undefined', () => {
    const all = buildSkillContext();
    const explicit = buildSkillContext(undefined);
    expect(explicit).toBe(all);
  });

  it('keeps @references directives as plain content without expansion', () => {
    const importedSkill: AgentSkill = {
      metadata: {
        name: 'Ref Tester',
        description: 'Skill with plain @references text.',
        version: '1.0.0',
        tags: [],
      },
      content: `Intro for references\n@references skill:excel, agent:Excel\nEnd of skill`,
    };

    setImportedSkills([importedSkill]);

    const context = buildSkillContext(['Ref Tester']);
    expect(context).toContain('@references skill:excel, agent:Excel');
    expect(context).not.toContain('Reference: skill:excel');
    expect(context).not.toContain('Reference: agent:Excel');
  });
});

describe('getSkills', () => {
  it('returns bundled skills', () => {
    const skills = getSkills();
    expect(skills.length).toBeGreaterThan(0);
    expect(skills.some(skill => skill.metadata.name === 'excel')).toBe(true);
  });

  it('each skill has metadata with a name', () => {
    for (const skill of getSkills()) {
      expect(skill.metadata.name).toBeTruthy();
    }
  });

  it('each skill has non-empty content', () => {
    for (const skill of getSkills()) {
      expect(skill.content.length).toBeGreaterThan(0);
    }
  });
});

describe('getSkill', () => {
  it('returns undefined for an unknown skill', () => {
    expect(getSkill('nonexistent-skill-xyz')).toBeUndefined();
  });

  it('returns bundled excel skill by name', () => {
    const skill = getSkill('excel');
    expect(skill).toBeDefined();
    expect(skill?.metadata.name).toBe('excel');
  });
});

describe('parseFrontmatter edge cases', () => {
  it('returns defaults when no frontmatter delimiters', () => {
    const { metadata, content } = parseFrontmatter('Just plain text');
    expect(metadata.name).toBe('unknown');
    expect(content).toBe('Just plain text');
  });

  it('returns defaults when closing delimiter is missing', () => {
    const { metadata } = parseFrontmatter('---\nname: test\n');
    expect(metadata.name).toBe('unknown');
  });

  it('parses simple key-value frontmatter', () => {
    const raw = `---
name: my-skill
description: A test skill
version: 1.0.0
---
Body content here`;
    const { metadata, content } = parseFrontmatter(raw);
    expect(metadata.name).toBe('my-skill');
    expect(metadata.description).toBe('A test skill');
    expect(metadata.version).toBe('1.0.0');
    expect(content).toBe('Body content here');
  });

  it('parses tags array', () => {
    const raw = `---
name: tagged
tags:
  - azure
  - excel
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata.tags).toEqual(['azure', 'excel']);
  });

  it('ignores non-standard references field', () => {
    const raw = `---
name: with-refs
references:
  - skill:excel
  - agent:Excel
---
Content`;
    const { metadata } = parseFrontmatter(raw);
    expect(metadata).not.toHaveProperty('references');
  });
});
