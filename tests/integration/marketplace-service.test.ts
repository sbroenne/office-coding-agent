// @vitest-environment node
/**
 * Real integration tests for marketplaceService.mjs.
 *
 * These tests create actual temp directories on disk — no mocking.
 * They verify the REAL logic for listing and removing marketplaces.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  listMarketplaces,
  removeMarketplace,
  repoCacheSlugs,
  OCA_MARKETPLACE_KEY,
  BUILTIN_KEYS,
} from '@/../src/marketplaceService.mjs';

// ─── Helpers ──────────────────────────────────────────────────────────────────

function tempDir(): string {
  const dir = join(tmpdir(), `oca-mkt-test-${randomUUID()}`);
  mkdirSync(dir, { recursive: true });
  return dir;
}

function makeConfig(marketplaces: Record<string, unknown>): string {
  return JSON.stringify({ marketplaces });
}

function makeCacheDir(baseDir: string, slug: string, manifestName: string | null = null): string {
  const dir = join(baseDir, slug);
  mkdirSync(dir, { recursive: true });
  if (manifestName !== null) {
    const pluginDir = join(dir, '.claude-plugin');
    mkdirSync(pluginDir, { recursive: true });
    writeFileSync(
      join(pluginDir, 'marketplace.json'),
      JSON.stringify({ name: manifestName, plugins: [] })
    );
  }
  return dir;
}

const cleanups: string[] = [];
afterEach(() => {
  for (const dir of cleanups.splice(0)) {
    try { rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ }
  }
});

// ─── repoCacheSlugs ───────────────────────────────────────────────────────────

describe('repoCacheSlugs', () => {
  it('produces owner-repo form', () => {
    expect(repoCacheSlugs('owner/repo')).toContain('owner-repo');
  });

  it('produces owner--repo form', () => {
    expect(repoCacheSlugs('owner/repo')).toContain('owner--repo');
  });

  it('produces fully slugified form for names with special chars', () => {
    expect(repoCacheSlugs('stbrnner_microsoft/SPT-IQ')).toContain('stbrnner-microsoft-spt-iq');
  });

  it('produces double-dash form observed in real CLI output', () => {
    expect(repoCacheSlugs('stbrnner_microsoft/SPT-IQ')).toContain('stbrnner_microsoft--SPT-IQ');
  });
});

// ─── listMarketplaces ─────────────────────────────────────────────────────────

describe('listMarketplaces', () => {
  it('registered marketplace always has non-null registeredKey (simple owner/repo)', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, 'sbroenne-my-plugins', 'My Plugins');
    writeFileSync(configPath, makeConfig({
      'my-plugins': { source: { source: 'github', repo: 'sbroenne/my-plugins' } },
    }));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.slug === 'sbroenne-my-plugins');
    expect(entry).toBeTruthy();
    expect(entry!.registeredKey).toBe('my-plugins');
  });

  it('registered marketplace with double-dash cache slug gets correct registeredKey', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, 'stbrnner_microsoft--SPT-IQ', 'SPT-IQ');
    writeFileSync(configPath, makeConfig({
      'spt-iq': { source: { source: 'github', repo: 'stbrnner_microsoft/SPT-IQ' } },
    }));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.slug === 'stbrnner_microsoft--SPT-IQ');
    expect(entry).toBeTruthy();
    expect(entry!.registeredKey).toBe('spt-iq');
  });

  it('cache-only dirs (not in config) are NOT listed', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, 'some-random-cache', null);
    writeFileSync(configPath, makeConfig({}));

    const list = listMarketplaces(cacheDir, configPath);
    expect(list.find(m => m.slug === 'some-random-cache')).toBeUndefined();
    expect(list).toHaveLength(0);
  });

  it('OCA marketplace entry has isOwn = true', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, 'sbroenne-office-coding-agent-plugins', 'office-coding-agent');
    writeFileSync(configPath, makeConfig({
      [OCA_MARKETPLACE_KEY]: { source: { source: 'github', repo: 'sbroenne/office-coding-agent-plugins' } },
    }));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.isOwn);
    expect(entry).toBeTruthy();
    expect(entry!.registeredKey).toBe(OCA_MARKETPLACE_KEY);
  });

  it('returns manifest name when available, falls back to config key', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, 'owner-repo', 'Pretty Name From Manifest');
    writeFileSync(configPath, makeConfig({
      'owner-repo': { source: { source: 'github', repo: 'owner/repo' } },
    }));

    const list = listMarketplaces(cacheDir, configPath);
    expect(list[0].name).toBe('Pretty Name From Manifest');
  });

  it('shows registered entry even when cache dir does not exist', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    writeFileSync(configPath, makeConfig({
      'no-cache-market': { source: { source: 'github', repo: 'owner/no-cache' } },
    }));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.registeredKey === 'no-cache-market');
    expect(entry).toBeTruthy();
    expect(entry!.registeredKey).toBe('no-cache-market');
  });

  it('returns empty list when config and cache both missing', () => {
    const root = tempDir(); cleanups.push(root);
    const list = listMarketplaces(join(root, 'cache'), join(root, 'config.json'));
    expect(list).toEqual([]);
  });
});

// ─── removeMarketplace ────────────────────────────────────────────────────────

describe('removeMarketplace', () => {
  it('refuses to remove OCA marketplace', () => {
    const result = removeMarketplace(OCA_MARKETPLACE_KEY);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/office-coding-agent/i);
  });

  it('refuses to remove built-in marketplace keys', () => {
    for (const key of BUILTIN_KEYS) {
      const result = removeMarketplace(key);
      expect(result.success).toBe(false);
    }
  });

  it('returns failure when CLI command fails (non-registered key)', () => {
    // The CLI will fail because the key is not registered; we expect a failure result
    const result = removeMarketplace('nonexistent-marketplace-key-zzz');
    expect(result.success).toBe(false);
  });
});
