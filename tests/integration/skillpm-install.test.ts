/**
 * Real integration test for skillpm skill package installation.
 *
 * This test actually runs `npm install skillpm-skill` into a temp directory
 * and verifies that the skill dirs are correctly discovered using the same
 * logic as copilotProxy.mjs:findSkillDirs.
 *
 * skillpm packages follow the spec: skills/<name>/SKILL.md inside the package.
 * See https://skillpm.dev for details.
 *
 * No mocks. This test hits the real npm registry and real filesystem.
 * It requires network access and npm in PATH.
 */

import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { mkdtemp, rm, readdir, access } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
/** On Windows, npm is npm.cmd */
const NPM_CMD = process.platform === 'win32' ? 'npm.cmd' : 'npm';
/** shell: true is required on Windows for .cmd executables */
const NPM_EXEC_OPTS = process.platform === 'win32' ? { shell: true } : {};

// ── Helper: same logic as copilotProxy.mjs ──────────────────────────────────

/** Find all skills/<name>/ dirs containing SKILL.md within a single package dir */
async function findSkillDirsInPackage(pkgDir: string): Promise<string[]> {
  const skillsRoot = join(pkgDir, 'skills');
  const result: string[] = [];
  let skillSubdirs;
  try {
    skillSubdirs = await readdir(skillsRoot, { withFileTypes: true });
  } catch {
    return result;
  }
  for (const sub of skillSubdirs) {
    if (sub.isDirectory()) {
      const skillDir = join(skillsRoot, sub.name);
      if (existsSync(join(skillDir, 'SKILL.md'))) {
        result.push(skillDir);
      }
    }
  }
  return result;
}

/** Scan node_modules for skillpm-compatible skill directories */
async function findSkillDirs(nodeModulesDir: string): Promise<string[]> {
  const skillDirs: string[] = [];
  let pkgEntries;
  try {
    pkgEntries = await readdir(nodeModulesDir, { withFileTypes: true });
  } catch {
    return skillDirs;
  }

  for (const entry of pkgEntries) {
    if (entry.isDirectory() && entry.name.startsWith('@')) {
      const scopeDir = join(nodeModulesDir, entry.name);
      let scopedEntries;
      try {
        scopedEntries = await readdir(scopeDir, { withFileTypes: true });
      } catch {
        continue;
      }
      for (const scopedEntry of scopedEntries) {
        if (scopedEntry.isDirectory()) {
          const pkgDir = join(scopeDir, scopedEntry.name);
          skillDirs.push(...(await findSkillDirsInPackage(pkgDir)));
        }
      }
    } else if (entry.isDirectory()) {
      const pkgDir = join(nodeModulesDir, entry.name);
      skillDirs.push(...(await findSkillDirsInPackage(pkgDir)));
    }
  }
  return skillDirs;
}

// ── Test setup ───────────────────────────────────────────────────────────────

let installDir: string;

beforeAll(async () => {
  // Create a fresh temp dir and install skillpm-skill (zero deps, fast install)
  installDir = await mkdtemp(join(tmpdir(), 'oca-skillpm-test-'));
  await execFileAsync(
    NPM_CMD,
    ['install', '--prefix', installDir, '--no-save', 'skillpm-skill'],
    { ...NPM_EXEC_OPTS, timeout: 90_000 },
  );
}, 120_000); // allow up to 120s for cold npm cache

afterAll(async () => {
  if (installDir) {
    await rm(installDir, { recursive: true, force: true });
  }
});

// ── Tests ────────────────────────────────────────────────────────────────────

describe('skillpm real install: skillpm-skill', () => {
  it('npm install succeeds and node_modules/skillpm-skill exists', async () => {
    const pkgDir = join(installDir, 'node_modules', 'skillpm-skill');
    expect(existsSync(pkgDir)).toBe(true);
  });

  it('package follows skillpm spec: skills/<name>/SKILL.md exists', async () => {
    const pkgDir = join(installDir, 'node_modules', 'skillpm-skill');
    const skillsDir = join(pkgDir, 'skills');
    expect(existsSync(skillsDir)).toBe(true);

    const subdirs = await readdir(skillsDir, { withFileTypes: true });
    const skillSubdirs = subdirs.filter(e => e.isDirectory());
    expect(skillSubdirs.length).toBeGreaterThan(0);

    for (const sub of skillSubdirs) {
      const skillMd = join(skillsDir, sub.name, 'SKILL.md');
      await expect(access(skillMd)).resolves.toBeUndefined();
    }
  });

  it('findSkillDirs discovers the skill directories', async () => {
    const nodeModulesDir = join(installDir, 'node_modules');
    const skillDirs = await findSkillDirs(nodeModulesDir);

    expect(skillDirs.length).toBeGreaterThan(0);
    // Every returned path should have a SKILL.md
    for (const dir of skillDirs) {
      expect(existsSync(join(dir, 'SKILL.md'))).toBe(true);
    }
  });

  it('findSkillDirs returns the skillpm skill dir specifically', async () => {
    const nodeModulesDir = join(installDir, 'node_modules');
    const skillDirs = await findSkillDirs(nodeModulesDir);

    // skillpm-skill publishes skills/skillpm/SKILL.md
    const skillpmDir = skillDirs.find(d => d.endsWith('skillpm'));
    expect(skillpmDir).toBeDefined();
    expect(existsSync(join(skillpmDir!, 'SKILL.md'))).toBe(true);
  });

  it('SKILL.md contains valid YAML frontmatter (name and description)', async () => {
    const nodeModulesDir = join(installDir, 'node_modules');
    const skillDirs = await findSkillDirs(nodeModulesDir);
    const skillpmDir = skillDirs.find(d => d.endsWith('skillpm'));
    expect(skillpmDir).toBeDefined();

    const { readFile } = await import('node:fs/promises');
    const content = await readFile(join(skillpmDir!, 'SKILL.md'), 'utf-8');

    // Frontmatter starts and ends with ---
    expect(content.trimStart().startsWith('---')).toBe(true);
    expect(content).toContain('name:');
    expect(content).toContain('description:');
  });
});

describe('skillpm findSkillDirs: scoped package layout', () => {
  it('handles scoped packages (@org/name) in node_modules', async () => {
    // Use the installed node_modules from previous test — verify scoped scanning
    // doesn't error even with no scoped packages present
    const nodeModulesDir = join(installDir, 'node_modules');
    // Should not throw even if there are no scoped packages
    const skillDirs = await findSkillDirs(nodeModulesDir);
    expect(Array.isArray(skillDirs)).toBe(true);
  });
});

describe('skillpm findSkillDirs: missing / empty directories', () => {
  it('returns empty array for non-existent node_modules', async () => {
    const skillDirs = await findSkillDirs('/nonexistent/path/node_modules');
    expect(skillDirs).toEqual([]);
  });

  it('returns empty array for a package with no skills/ directory', async () => {
    // installDir itself has no skills/ folder
    const dirs = await findSkillDirsInPackage(installDir);
    expect(dirs).toEqual([]);
  });
});
