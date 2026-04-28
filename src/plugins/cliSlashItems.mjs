import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const DEFAULT_INSTALLED_PLUGINS_DIR = path.join(os.homedir(), '.copilot', 'installed-plugins');
const DEFAULT_WORKSPACE_PROMPTS_DIR = path.join(process.cwd(), '.github', 'prompts');

function parseFrontmatter(markdown) {
  if (!markdown.startsWith('---')) return {};
  const end = markdown.indexOf('\n---', 3);
  if (end === -1) return {};
  const yaml = markdown.slice(3, end).trim();
  const result = {};
  for (const line of yaml.split(/\r?\n/)) {
    const match = line.match(/^([A-Za-z0-9_-]+):\s*(.+)$/);
    if (!match) continue;
    result[match[1]] = match[2].replace(/^['"]|['"]$/g, '').trim();
  }
  return result;
}

async function walk(dir, visitor) {
  let entries;
  try {
    entries = await fs.promises.readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await walk(fullPath, visitor);
    } else if (entry.isFile()) {
      await visitor(fullPath);
    }
  }
}

function pluginNameFor(filePath, installedPluginsDir) {
  const relative = path.relative(installedPluginsDir, filePath);
  const parts = relative.split(path.sep);
  return parts.length >= 2 ? `${parts[1]}@${parts[0]}` : undefined;
}

async function readSkill(filePath, installedPluginsDir) {
  const markdown = await fs.promises.readFile(filePath, 'utf8');
  const metadata = parseFrontmatter(markdown);
  const fallbackName = path.basename(path.dirname(filePath));
  const bodyStart = markdown.startsWith('---') ? markdown.indexOf('\n---', 3) : -1;
  const body = bodyStart === -1 ? markdown : markdown.slice(bodyStart + 4);
  const firstBodyLine = body
    .split(/\r?\n/)
    .map(line => line.trim())
    .find(Boolean);

  return {
    type: 'skill',
    name: String(metadata.name ?? fallbackName),
    description: String(metadata.description ?? firstBodyLine ?? ''),
    plugin: pluginNameFor(filePath, installedPluginsDir),
  };
}

async function readPrompt(filePath, source) {
  const markdown = await fs.promises.readFile(filePath, 'utf8');
  const metadata = parseFrontmatter(markdown);
  const fallbackName = path.basename(filePath).replace(/\.prompt\.md$/i, '');
  const bodyStart = markdown.startsWith('---') ? markdown.indexOf('\n---', 3) : -1;
  const body = bodyStart === -1 ? markdown : markdown.slice(bodyStart + 4);
  const firstBodyLine = body
    .split(/\r?\n/)
    .map(line => line.trim())
    .find(Boolean);

  return {
    type: 'prompt',
    name: String(metadata.name ?? fallbackName),
    description: String(metadata.description ?? firstBodyLine ?? ''),
    source,
  };
}

function isPromptFile(filePath) {
  return path.basename(filePath).toLowerCase().endsWith('.prompt.md');
}

export async function getCliSlashItems(options = {}) {
  const installedPluginsDir = options.installedPluginsDir ?? DEFAULT_INSTALLED_PLUGINS_DIR;
  const workspacePromptsDir = options.workspacePromptsDir ?? DEFAULT_WORKSPACE_PROMPTS_DIR;
  const skills = [];
  const prompts = [];

  await walk(installedPluginsDir, async filePath => {
    const base = path.basename(filePath).toLowerCase();
    if (base === 'skill.md') {
      skills.push(await readSkill(filePath, installedPluginsDir));
      return;
    }
    if (isPromptFile(filePath)) {
      prompts.push(await readPrompt(filePath, pluginNameFor(filePath, installedPluginsDir)));
    }
  });

  await walk(workspacePromptsDir, async filePath => {
    if (isPromptFile(filePath)) {
      prompts.push(await readPrompt(filePath, 'workspace'));
    }
  });

  const byName = (a, b) => a.name.localeCompare(b.name);
  return {
    skills: skills.sort(byName),
    prompts: prompts.sort(byName),
  };
}
