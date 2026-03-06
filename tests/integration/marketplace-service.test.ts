// @vitest-environment node
/**
 * Real integration tests for marketplaceService.mjs.
 * Uses actual temp directories on disk — no mocking.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { mkdirSync, writeFileSync, existsSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  listMarketplaces,
  removeMarketplace,
  repoCacheSlugs,
  OCA_MARKETPLACE_KEY,
  OCA_MARKETPLACE_SLUG,
  BUILTIN_KEYS,
  BUILTIN_SLUGS,
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
  it('registered marketplace has non-null registeredKey', () => {
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

  it('cache-only dir (not in config) IS listed with registeredKey=null', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, 'some-cache-only-marketplace', null);
    writeFileSync(configPath, makeConfig({}));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.slug === 'some-cache-only-marketplace');
    expect(entry).toBeTruthy();
    expect(entry!.registeredKey).toBeNull();
  });

  it('OCA marketplace entry has isOwn = true', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, OCA_MARKETPLACE_SLUG, 'office-coding-agent');
    writeFileSync(configPath, makeConfig({
      [OCA_MARKETPLACE_KEY]: { source: { source: 'github', repo: 'sbroenne/office-coding-agent-plugins' } },
    }));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.isOwn);
    expect(entry).toBeTruthy();
    expect(entry!.registeredKey).toBe(OCA_MARKETPLACE_KEY);
  });

  it('built-in cache-only dir is listed with isBuiltIn=true', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    const configPath = join(root, 'config.json');

    makeCacheDir(cacheDir, BUILTIN_SLUGS[0], null);
    writeFileSync(configPath, makeConfig({}));

    const list = listMarketplaces(cacheDir, configPath);
    const entry = list.find(m => m.slug === BUILTIN_SLUGS[0]);
    expect(entry).toBeTruthy();
    expect(entry!.isBuiltIn).toBe(true);
    expect(entry!.registeredKey).toBeNull();
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
  it('removes cache directory for a cache-only marketplace', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');

    makeCacheDir(cacheDir, 'my-cache-slug');
    expect(existsSync(join(cacheDir, 'my-cache-slug'))).toBe(true);

    const result = removeMarketplace(cacheDir, 'my-cache-slug', null);
    expect(result.success).toBe(true);
    expect(existsSync(join(cacheDir, 'my-cache-slug'))).toBe(false);
  });

  it('does not throw when cache dir does not exist', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    mkdirSync(cacheDir);

    const result = removeMarketplace(cacheDir, 'nonexistent-slug', null);
    expect(result.success).toBe(true);
  });

  it('refuses to remove OCA marketplace (by slug)', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    makeCacheDir(cacheDir, OCA_MARKETPLACE_SLUG);

    const result = removeMarketplace(cacheDir, OCA_MARKETPLACE_SLUG, null);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/office-coding-agent/i);
    expect(existsSync(join(cacheDir, OCA_MARKETPLACE_SLUG))).toBe(true);
  });

  it('refuses to remove OCA marketplace (by registeredKey)', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');

    const result = removeMarketplace(cacheDir, 'some-slug', OCA_MARKETPLACE_KEY);
    expect(result.success).toBe(false);
  });

  it('refuses to remove built-in marketplace (by slug)', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');
    makeCacheDir(cacheDir, BUILTIN_SLUGS[0]);

    const result = removeMarketplace(cacheDir, BUILTIN_SLUGS[0], null);
    expect(result.success).toBe(false);
  });

  it('refuses to remove built-in marketplace (by registeredKey)', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');

    for (const key of BUILTIN_KEYS) {
      const result = removeMarketplace(cacheDir, 'some-slug', key);
      expect(result.success).toBe(false);
    }
  });

  it('returns failure when CLI command fails for a registered key', () => {
    const root = tempDir(); cleanups.push(root);
    const cacheDir = join(root, 'cache');

    const result = removeMarketplace(cacheDir, 'nonexistent-key-zzz', 'nonexistent-key-zzz');
    expect(result.success).toBe(false);
  });
});
