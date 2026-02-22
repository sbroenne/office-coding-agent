import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import {
  parseAgentsZipFile,
  parseSkillsZipFile,
  parseAgentMarkdownFile,
  parseSkillMarkdownFile,
} from '@/services/extensions/zipImportService';

async function createZipFile(name: string, entries: Record<string, string>): Promise<File> {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(entries)) {
    zip.file(path, content);
  }
  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  return new File([buffer], name, { type: 'application/zip' });
}

describe('zipImportService', () => {
  it('parses skills from skills/ folder in ZIP', async () => {
    const file = await createZipFile('skills.zip', {
      'skills/custom-skill.md': `---
name: Custom Skill
description: Imported skill
version: 1.0.0
---

Skill body`,
    });

    const skills = await parseSkillsZipFile(file);

    expect(skills).toHaveLength(1);
    expect(skills[0].metadata.name).toBe('Custom Skill');
    expect(skills[0].content).toContain('Skill body');
  });

  it('fails skills import when no skills/ markdown files exist', async () => {
    const file = await createZipFile('skills.zip', {
      'notes/readme.md': 'not a skill file',
    });

    await expect(parseSkillsZipFile(file)).rejects.toThrow(
      "No markdown files found in 'skills/' folder."
    );
  });

  it('parses agents from agents/ folder in ZIP', async () => {
    const file = await createZipFile('agents.zip', {
      'agents/custom-agent.md': `---
name: Custom Agent
description: Imported agent
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---

Agent instructions`,
    });

    const agents = await parseAgentsZipFile(file);

    expect(agents).toHaveLength(1);
    expect(agents[0].metadata.name).toBe('Custom Agent');
    expect(agents[0].metadata.hosts).toEqual(['excel']);
  });

  it('fails agents import when host list is missing/invalid', async () => {
    const file = await createZipFile('agents.zip', {
      'agents/custom-agent.md': `---
name: Custom Agent
description: Imported agent
version: 1.0.0
---

Agent instructions`,
    });

    await expect(parseAgentsZipFile(file)).rejects.toThrow(
      "Agent file 'agents/custom-agent.md' must include at least one supported host."
    );
  });
});

describe('parseAgentMarkdownFile', () => {
  function makeMdFile(content: string, name = 'agent.md'): File {
    return new File([content], name, { type: 'text/markdown' });
  }

  it('parses a valid agent .md file', async () => {
    const file = makeMdFile(`---
name: My Agent
description: A custom agent
version: 1.0.0
hosts: [excel]
defaultForHosts: [excel]
---

Agent instructions here.`);
    const agent = await parseAgentMarkdownFile(file);
    expect(agent.metadata.name).toBe('My Agent');
    expect(agent.metadata.hosts).toEqual(['excel']);
    expect(agent.instructions).toContain('Agent instructions here.');
  });

  it('parses tools and mcpServers fields', async () => {
    const file = makeMdFile(`---
name: Scoped Agent
description: desc
version: 1.0.0
hosts: [excel]
defaultForHosts: []
tools: [create_chart, format_range]
mcpServers: [my-server]
---
Instructions`);
    const agent = await parseAgentMarkdownFile(file);
    expect(agent.metadata.tools).toEqual(['create_chart', 'format_range']);
    expect(agent.metadata.mcpServers).toEqual(['my-server']);
  });

  it('rejects a non-.md file', async () => {
    const file = new File(['content'], 'agent.zip');
    await expect(parseAgentMarkdownFile(file)).rejects.toThrow('.md');
  });

  it('rejects when name is missing', async () => {
    const file = makeMdFile(`---
description: no name here
version: 1.0.0
hosts: [excel]
---
Instructions`);
    await expect(parseAgentMarkdownFile(file)).rejects.toThrow('name');
  });

  it('rejects when hosts are missing', async () => {
    const file = makeMdFile(`---
name: No Hosts Agent
description: desc
version: 1.0.0
---
Instructions`);
    await expect(parseAgentMarkdownFile(file)).rejects.toThrow('host');
  });
});

describe('parseSkillMarkdownFile', () => {
  function makeMdFile(content: string, name = 'skill.md'): File {
    return new File([content], name, { type: 'text/markdown' });
  }

  it('parses a valid skill .md file', async () => {
    const file = makeMdFile(`---
name: My Skill
description: A custom skill
version: 1.0.0
---

Skill content here.`);
    const skill = await parseSkillMarkdownFile(file);
    expect(skill.metadata.name).toBe('My Skill');
    expect(skill.content).toContain('Skill content here.');
  });

  it('rejects a non-.md file', async () => {
    const file = new File(['content'], 'skill.zip');
    await expect(parseSkillMarkdownFile(file)).rejects.toThrow('.md');
  });

  it('rejects when name is missing', async () => {
    const file = makeMdFile(`---
description: no name
version: 1.0.0
---
Content`);
    await expect(parseSkillMarkdownFile(file)).rejects.toThrow('name');
  });
});

// ─── File size limits ─────────────────────────────────────────────────────────

describe('file size limits', () => {
  it('parseAgentMarkdownFile rejects a file larger than 1 MB', async () => {
    // 1 MB + 1 byte of valid-looking content
    const content = 'x'.repeat(1024 * 1024 + 1);
    const file = new File([content], 'large.md', { type: 'text/markdown' });

    await expect(parseAgentMarkdownFile(file)).rejects.toThrow(/too large/i);
  });

  it('parseSkillMarkdownFile rejects a file larger than 1 MB', async () => {
    const content = 'x'.repeat(1024 * 1024 + 1);
    const file = new File([content], 'large.md', { type: 'text/markdown' });

    await expect(parseSkillMarkdownFile(file)).rejects.toThrow(/too large/i);
  });

  it('parseAgentsZipFile rejects a ZIP larger than 5 MB', async () => {
    // The size check happens before ZIP parsing, so any large binary blob works
    const buf = new Uint8Array(5 * 1024 * 1024 + 1).fill(0x50); // 'P' — not valid ZIP
    const file = new File([buf], 'large.zip', { type: 'application/zip' });

    await expect(parseAgentsZipFile(file)).rejects.toThrow(/too large/i);
  });

  it('parseSkillsZipFile rejects a ZIP larger than 5 MB', async () => {
    const buf = new Uint8Array(5 * 1024 * 1024 + 1).fill(0x50);
    const file = new File([buf], 'large.zip', { type: 'application/zip' });

    await expect(parseSkillsZipFile(file)).rejects.toThrow(/too large/i);
  });
});


