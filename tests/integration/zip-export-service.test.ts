import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  slugify,
  agentToMarkdown,
  buildAgentsZip,
  buildSkillsZip,
} from '@/services/extensions/zipExportService';
import { skillToMarkdown } from '@/services/skills';
import { parseAgentFrontmatter } from '@/services/agents';
import { parseFrontmatter } from '@/services/skills';
import { parseAgentsZipFile, parseSkillsZipFile } from '@/services/extensions/zipImportService';
import type { AgentConfig } from '@/types/agent';
import type { AgentSkill } from '@/types/skill';

// ── Test fixtures ─────────────────────────────────────────────────────────

const testAgent: AgentConfig = {
  metadata: {
    name: 'My Test Agent',
    description: 'A test agent',
    version: '2.0.0',
    hosts: ['excel', 'word'],
    defaultForHosts: ['excel'],
    tools: ['create_chart', 'format_range'],
    mcpServers: ['my-server'],
  },
  instructions: 'Do the thing.\n\nSecond paragraph.',
};

const testAgentMinimal: AgentConfig = {
  metadata: {
    name: 'Simple Agent',
    description: 'Simple',
    version: '1.0.0',
    hosts: ['excel'],
    defaultForHosts: [],
  },
  instructions: 'Keep it simple.',
};

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

// ── agentToMarkdown ────────────────────────────────────────────────────────

describe('agentToMarkdown', () => {
  it('produces valid frontmatter that round-trips', () => {
    const md = agentToMarkdown(testAgent);
    const parsed = parseAgentFrontmatter(md);
    expect(parsed.metadata.name).toBe(testAgent.metadata.name);
    expect(parsed.metadata.description).toBe(testAgent.metadata.description);
    expect(parsed.metadata.version).toBe(testAgent.metadata.version);
    expect(parsed.metadata.hosts).toEqual(testAgent.metadata.hosts);
    expect(parsed.metadata.defaultForHosts).toEqual(testAgent.metadata.defaultForHosts);
    expect(parsed.metadata.tools).toEqual(testAgent.metadata.tools);
    expect(parsed.metadata.mcpServers).toEqual(testAgent.metadata.mcpServers);
    expect(parsed.instructions).toContain('Do the thing.');
  });

  it('omits tools line when tools is undefined', () => {
    const md = agentToMarkdown(testAgentMinimal);
    expect(md).not.toContain('tools:');
    expect(md).not.toContain('mcpServers:');
  });

  it('includes instructions after the closing ---', () => {
    const md = agentToMarkdown(testAgent);
    const parts = md.split('---');
    // [0] = '', [1] = frontmatter, [2] = instructions
    expect(parts.length).toBeGreaterThanOrEqual(3);
    expect(parts[2]).toContain('Do the thing.');
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

// ── buildAgentsZip round-trip ─────────────────────────────────────────────

describe('buildAgentsZip', () => {
  it('builds a blob that parseAgentsZipFile can read back', async () => {
    const blob = await buildAgentsZip([testAgent, testAgentMinimal]);
    const file = new File([blob], 'agents.zip', { type: 'application/zip' });
    const agents = await parseAgentsZipFile(file);

    expect(agents).toHaveLength(2);
    const names = agents.map(a => a.metadata.name);
    expect(names).toContain(testAgent.metadata.name);
    expect(names).toContain(testAgentMinimal.metadata.name);
  });

  it('round-tripped agent preserves tools and mcpServers', async () => {
    const blob = await buildAgentsZip([testAgent]);
    const file = new File([blob], 'agents.zip', { type: 'application/zip' });
    const [agent] = await parseAgentsZipFile(file);

    expect(agent.metadata.tools).toEqual(testAgent.metadata.tools);
    expect(agent.metadata.mcpServers).toEqual(testAgent.metadata.mcpServers);
  });

  it('places files under agents/ in the ZIP', async () => {
    const blob = await buildAgentsZip([testAgent]);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const paths = Object.keys(zip.files).filter(p => !zip.files[p].dir);
    expect(paths.length).toBeGreaterThan(0);
    expect(paths.every(p => p.startsWith('agents/'))).toBe(true);
    expect(paths[0]).toMatch(/\.md$/);
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
