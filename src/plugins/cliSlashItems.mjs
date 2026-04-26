import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const DEFAULT_INSTALLED_PLUGINS_DIR = path.join(os.homedir(), '.copilot', 'installed-plugins');

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

async function readMarkdownItem(filePath, type, installedPluginsDir) {
  const markdown = await fs.promises.readFile(filePath, 'utf8');
  const metadata = parseFrontmatter(markdown);
  const fallbackName =
    type === 'skill' && path.basename(filePath).toLowerCase() === 'skill.md'
      ? path.basename(path.dirname(filePath))
      : path.basename(filePath).replace(/\.prompt\.md$|\.md$/i, '');
  const bodyStart = markdown.startsWith('---') ? markdown.indexOf('\n---', 3) : -1;
  const body = bodyStart === -1 ? markdown : markdown.slice(bodyStart + 4);
  const firstBodyLine = body
    .split(/\r?\n/)
    .map(line => line.trim())
    .find(Boolean);

  return {
    type,
    name: String(metadata.name ?? fallbackName),
    description: String(metadata.description ?? firstBodyLine ?? ''),
    plugin: pluginNameFor(filePath, installedPluginsDir),
  };
}

function isPromptFile(filePath) {
  const lower = filePath.toLowerCase();
  return (
    lower.endsWith('.prompt.md') ||
    lower.endsWith(path.join('prompts', 'prompt.md').toLowerCase()) ||
    lower.includes(`${path.sep}prompts${path.sep}`.toLowerCase()) && lower.endsWith('.md')
  );
}

export async function getCliSlashItems(options = {}) {
  const installedPluginsDir = options.installedPluginsDir ?? DEFAULT_INSTALLED_PLUGINS_DIR;
  const skills = [];
  const prompts = [];

  await walk(installedPluginsDir, async filePath => {
    const base = path.basename(filePath).toLowerCase();
    if (base === 'skill.md') {
      skills.push(await readMarkdownItem(filePath, 'skill', installedPluginsDir));
      return;
    }
    if (isPromptFile(filePath)) {
      prompts.push(await readMarkdownItem(filePath, 'prompt', installedPluginsDir));
    }
  });

  const byName = (a, b) => a.name.localeCompare(b.name);
  return {
    skills: skills.sort(byName),
    prompts: prompts.sort(byName),
  };
}
