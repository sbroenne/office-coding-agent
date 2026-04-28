import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { parseSkillsZipFile, parseSkillMarkdownFile } from '@/services/extensions/zipImportService';

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
  it('parseSkillMarkdownFile rejects a file larger than 1 MB', async () => {
    const content = 'x'.repeat(1024 * 1024 + 1);
    const file = new File([content], 'large.md', { type: 'text/markdown' });

    await expect(parseSkillMarkdownFile(file)).rejects.toThrow(/too large/i);
  });

  it('parseSkillsZipFile rejects a ZIP larger than 5 MB', async () => {
    const buf = new Uint8Array(5 * 1024 * 1024 + 1).fill(0x50);
    const file = new File([buf], 'large.zip', { type: 'application/zip' });

    await expect(parseSkillsZipFile(file)).rejects.toThrow(/too large/i);
  });
});


