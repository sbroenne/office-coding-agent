import JSZip from 'jszip';
import { parseFrontmatter } from '@/services/skills';
import { parseAgentFrontmatter } from '@/services/agents';
import type { AgentSkill } from '@/types/skill';
import type { AgentConfig } from '@/types/agent';

const MAX_ZIP_BYTES = 5 * 1024 * 1024;
const MAX_TOTAL_MARKDOWN_BYTES = 2 * 1024 * 1024;

interface ZipMarkdownEntry {
  path: string;
  content: string;
}

function normalizeZipPath(path: string): string {
  return path.replace(/\\/g, '/').replace(/^\/+/, '');
}

function validateZipPath(path: string): void {
  if (!path || path.startsWith('/') || path.includes('..')) {
    throw new Error(`Invalid zip entry path: ${path}`);
  }
}

async function getMarkdownFilesFromFolder(file: File, folderName: 'skills' | 'agents') {
  if (file.size > MAX_ZIP_BYTES) {
    throw new Error(`ZIP file is too large. Maximum size is ${MAX_ZIP_BYTES / (1024 * 1024)}MB.`);
  }

  const zip = await JSZip.loadAsync(file);
  const entries: ZipMarkdownEntry[] = [];
  let totalMarkdownBytes = 0;

  for (const [rawPath, zipEntry] of Object.entries(zip.files)) {
    if (zipEntry.dir) continue;

    const path = normalizeZipPath(rawPath);
    validateZipPath(path);

    if (!path.toLowerCase().startsWith(`${folderName}/`)) continue;
    if (!path.toLowerCase().endsWith('.md')) continue;

    const content = await zipEntry.async('string');
    totalMarkdownBytes += content.length;

    if (totalMarkdownBytes > MAX_TOTAL_MARKDOWN_BYTES) {
      throw new Error('ZIP markdown content is too large after extraction.');
    }

    entries.push({ path, content });
  }

  if (entries.length === 0) {
    throw new Error(`No markdown files found in '${folderName}/' folder.`);
  }

  return entries;
}

export async function parseSkillsZipFile(file: File): Promise<AgentSkill[]> {
  const entries = await getMarkdownFilesFromFolder(file, 'skills');

  const skills = entries.map(entry => {
    const parsed = parseFrontmatter(entry.content);
    const name = parsed.metadata.name.trim();

    if (!name || name === 'unknown') {
      throw new Error(`Skill file '${entry.path}' is missing a valid frontmatter name.`);
    }

    return parsed;
  });

  return skills;
}

export async function parseAgentsZipFile(file: File): Promise<AgentConfig[]> {
  const entries = await getMarkdownFilesFromFolder(file, 'agents');

  const agents = entries.map(entry => {
    const parsed = parseAgentFrontmatter(entry.content);
    const name = parsed.metadata.name.trim();

    if (!name || name === 'unknown') {
      throw new Error(`Agent file '${entry.path}' is missing a valid frontmatter name.`);
    }

    if (parsed.metadata.hosts.length === 0) {
      throw new Error(`Agent file '${entry.path}' must include at least one supported host.`);
    }

    return parsed;
  });

  return agents;
}
