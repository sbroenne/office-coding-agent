import JSZip from 'jszip';
import type { AgentSkill } from '@/types/skill';
import { skillToMarkdown } from '@/services/skills';

/** Convert a display name to a safe lowercase filename slug. */
export function slugify(name: string): string {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

/** Trigger a browser file download with the given blob and filename. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/** Download a single skill as a `.md` file. */
export function downloadSkill(skill: AgentSkill): void {
  const content = skillToMarkdown(skill);
  const blob = new Blob([content], { type: 'text/markdown' });
  downloadBlob(blob, `${slugify(skill.metadata.name)}.md`);
}

/** Build a ZIP containing all given skills under a `skills/` folder. */
export async function buildSkillsZip(skills: AgentSkill[]): Promise<Blob> {
  const zip = new JSZip();
  for (const skill of skills) {
    const filename = `${slugify(skill.metadata.name)}.md`;
    zip.file(`skills/${filename}`, skillToMarkdown(skill));
  }
  return zip.generateAsync({ type: 'blob' });
}

/** Build and download all skills as `skills.zip`. */
export async function downloadSkillsZip(skills: AgentSkill[]): Promise<void> {
  const blob = await buildSkillsZip(skills);
  downloadBlob(blob, 'skills.zip');
}
