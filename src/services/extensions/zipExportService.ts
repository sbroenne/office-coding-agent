import JSZip from 'jszip';
import type { AgentConfig } from '@/types/agent';
import type { AgentSkill } from '@/types/skill';
import { skillToMarkdown } from '@/services/skills';

/** Convert a display name to a safe lowercase filename slug. */
export function slugify(name: string): string {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

/** Serialize an agent back to its YAML-frontmatter markdown format. */
export function agentToMarkdown(agent: AgentConfig): string {
  const { metadata, instructions } = agent;
  const lines: string[] = ['---'];
  lines.push(`name: ${metadata.name}`);
  lines.push(`description: ${metadata.description}`);
  lines.push(`version: ${metadata.version}`);
  lines.push(`hosts: [${metadata.hosts.join(', ')}]`);
  lines.push(`defaultForHosts: [${metadata.defaultForHosts.join(', ')}]`);
  if (metadata.tools && metadata.tools.length > 0) {
    lines.push(`tools: [${metadata.tools.join(', ')}]`);
  }
  if (metadata.mcpServers && metadata.mcpServers.length > 0) {
    lines.push(`mcpServers: [${metadata.mcpServers.join(', ')}]`);
  }
  lines.push('---');
  lines.push('');
  lines.push(instructions);
  return lines.join('\n');
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

/** Download a single agent as a `.md` file. */
export function downloadAgent(agent: AgentConfig): void {
  const content = agentToMarkdown(agent);
  const blob = new Blob([content], { type: 'text/markdown' });
  downloadBlob(blob, `${slugify(agent.metadata.name)}.md`);
}

/** Download a single skill as a `.md` file. */
export function downloadSkill(skill: AgentSkill): void {
  const content = skillToMarkdown(skill);
  const blob = new Blob([content], { type: 'text/markdown' });
  downloadBlob(blob, `${slugify(skill.metadata.name)}.md`);
}

/** Build a ZIP containing all given agents under an `agents/` folder. */
export async function buildAgentsZip(agents: AgentConfig[]): Promise<Blob> {
  const zip = new JSZip();
  for (const agent of agents) {
    const filename = `${slugify(agent.metadata.name)}.md`;
    zip.file(`agents/${filename}`, agentToMarkdown(agent));
  }
  return zip.generateAsync({ type: 'blob' });
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

/** Build and download all agents as `agents.zip`. */
export async function downloadAgentsZip(agents: AgentConfig[]): Promise<void> {
  const blob = await buildAgentsZip(agents);
  downloadBlob(blob, 'agents.zip');
}

/** Build and download all skills as `skills.zip`. */
export async function downloadSkillsZip(skills: AgentSkill[]): Promise<void> {
  const blob = await buildSkillsZip(skills);
  downloadBlob(blob, 'skills.zip');
}
