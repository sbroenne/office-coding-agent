import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { parseAgentsZipFile, parseSkillsZipFile } from '@/services/extensions/zipImportService';

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
