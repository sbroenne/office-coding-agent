/**
 * Unit tests for buildSkillContext and related skill functions.
 *
 * These exercise the real `skillService` module which imports
 * bundled `.md` skill files via the rawMarkdownPlugin in vitest.config.ts.
 */

import { describe, it, expect } from 'vitest';
import {
  buildSkillContext,
  getSkills,
  getSkill,
  parseFrontmatter,
} from '@/services/skills/skillService';

describe('buildSkillContext', () => {
  it('returns empty string when no bundled skills exist', () => {
    const ctx = buildSkillContext();
    expect(ctx).toBe('');
  });

  it('returns empty string when activeNames provided but no skills are bundled', () => {
    const ctx = buildSkillContext(['any-skill']);
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
});

describe('getSkills', () => {
  it('returns empty list when no bundled skills are configured', () => {
    const skills = getSkills();
    expect(skills).toEqual([]);
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

  it('returns undefined when no skills are bundled', () => {
    expect(getSkill('any-skill')).toBeUndefined();
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
});
