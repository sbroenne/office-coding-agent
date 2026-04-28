import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { slugify, buildSkillsZip } from '@/services/extensions/zipExportService';
import { skillToMarkdown } from '@/services/skills';
import { parseFrontmatter } from '@/services/skills';
import { parseSkillsZipFile } from '@/services/extensions/zipImportService';
import type { AgentSkill } from '@/types/skill';

// ── Test fixtures ─────────────────────────────────────────────────────────

const testSkill: AgentSkill = {
  metadata: { name: 'My Test Skill', description: 'A test skill', version: '1.1.0', tags: [], hosts: [] },
  content: 'Skill instructions here.',
};

// ── slugify ───────────────────────────────────────────────────────────────

describe('slugify', () => {
  it('lowercases and replaces spaces with hyphens', () => {
    expect(slugify('My Test Agent')).toBe('my-test-agent');
  });

  it('removes special characters', () => {
    expect(slugify('Agent: v2.0!')).toBe('agent-v2-0');
  });

  it('strips leading/trailing hyphens', () => {
    expect(slugify('  Hello  ')).toBe('hello');
  });

  it('collapses multiple non-alphanumeric chars', () => {
    expect(slugify('A  --  B')).toBe('a-b');
  });
});

// ── skillToMarkdown ────────────────────────────────────────────────────────

describe('skillToMarkdown', () => {
  it('produces valid frontmatter that round-trips', () => {
    const md = skillToMarkdown(testSkill);
    const parsed = parseFrontmatter(md);
    expect(parsed.metadata.name).toBe(testSkill.metadata.name);
    expect(parsed.metadata.description).toBe(testSkill.metadata.description);
    expect(parsed.metadata.version).toBe(testSkill.metadata.version);
    expect(parsed.content).toContain('Skill instructions here.');
  });
});

// ── buildSkillsZip round-trip ─────────────────────────────────────────────

describe('buildSkillsZip', () => {
  it('builds a blob that parseSkillsZipFile can read back', async () => {
    const blob = await buildSkillsZip([testSkill]);
    const file = new File([blob], 'skills.zip', { type: 'application/zip' });
    const skills = await parseSkillsZipFile(file);

    expect(skills).toHaveLength(1);
    expect(skills[0].metadata.name).toBe(testSkill.metadata.name);
    expect(skills[0].content).toContain('Skill instructions here.');
  });

  it('places files under skills/ in the ZIP', async () => {
    const blob = await buildSkillsZip([testSkill]);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const paths = Object.keys(zip.files).filter(p => !zip.files[p].dir);
    expect(paths.length).toBeGreaterThan(0);
    expect(paths.every(p => p.startsWith('skills/'))).toBe(true);
  });
});
