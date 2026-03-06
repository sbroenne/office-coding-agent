/**
 * marketplaceService.mjs
 *
 * Testable helper functions for marketplace listing and removal.
 * Extracted from server.mjs so the logic can be tested without Express.
 */

import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';

// ─── Constants ────────────────────────────────────────────────────────────────

export const OCA_MARKETPLACE_KEY = 'office-coding-agent';
export const OCA_MARKETPLACE_SLUG = 'sbroenne-office-coding-agent-plugins';
export const BUILTIN_KEYS = ['copilot-plugins', 'awesome-copilot'];
export const BUILTIN_SLUGS = ['github-copilot-plugins', 'github-awesome-copilot'];

// ─── File system helpers ──────────────────────────────────────────────────────

export function readConfig(configPath) {
  if (!fs.existsSync(configPath)) return {};
  try {
    return JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  } catch {
    return {};
  }
}

/**
 * Find the marketplace.json manifest inside a cache directory.
 * The CLI clones the GitHub repo so the manifest may live in various locations.
 */
export function findMarketplaceManifest(cacheDir) {
  const candidates = [
    path.join(cacheDir, '.claude-plugin', 'marketplace.json'),
    path.join(cacheDir, '.github', 'plugin', 'marketplace.json'),
    path.join(cacheDir, 'marketplace.json'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      try {
        return JSON.parse(fs.readFileSync(p, 'utf-8'));
      } catch {
        continue;
      }
    }
  }
  return null;
}

// ─── Slug helpers ─────────────────────────────────────────────────────────────

/**
 * Convert an "owner/repo" string to all plausible cache-dir name forms the
 * Copilot CLI might use.
 */
export function repoCacheSlugs(repo) {
  const slugs = new Set();
  slugs.add(repo.replace('/', '-'));
  slugs.add(repo.replace('/', '--'));
  slugs.add(repo.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, ''));
  const [owner, repoName] = repo.split('/');
  if (owner && repoName) {
    const repoSlug = repoName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
    slugs.add(`${owner}-${repoSlug}`);
    slugs.add(`${owner}--${repoSlug}`);
  }
  return [...slugs];
}

// ─── listMarketplaces ─────────────────────────────────────────────────────────

/**
 * List ALL known marketplaces — registered (from config.json) PLUS cache-only
 * directories the CLI has cloned but the user never explicitly registered.
 *
 * Registered entries  → registeredKey = config key (non-null)
 * Cache-only entries  → registeredKey = null  (removed by fs.rmSync)
 *
 * @param {string} cacheDir   - path to ~/.copilot/marketplace-cache
 * @param {string} configPath - path to ~/.copilot/config.json
 */
export function listMarketplaces(cacheDir, configPath) {
  const config = readConfig(configPath);
  const configMarketplaces = config.marketplaces || {};
  const result = [];
  const coveredSlugs = new Set();

  // ── Step 1: registered entries (config.json) ──────────────────────────────
  for (const [key, value] of Object.entries(configMarketplaces)) {
    const repo = value?.source?.repo ?? null;
    const isBuiltIn = BUILTIN_KEYS.includes(key);
    const isOwn = key === OCA_MARKETPLACE_KEY;

    let slug = null;
    let manifest = null;
    let pluginCount = 0;

    if (fs.existsSync(cacheDir) && repo) {
      const candidateSlugs = repoCacheSlugs(repo);
      const entries = fs.readdirSync(cacheDir, { withFileTypes: true });
      for (const entry of entries) {
        if (!entry.isDirectory()) continue;
        if (candidateSlugs.includes(entry.name) || entry.name === key) {
          slug = entry.name;
          manifest = findMarketplaceManifest(path.join(cacheDir, entry.name));
          pluginCount = Array.isArray(manifest?.plugins) ? manifest.plugins.length : 0;
          coveredSlugs.add(entry.name);
          break;
        }
      }
    }

    result.push({
      slug: slug ?? key,
      name: manifest?.name ?? key,
      source: repo ?? key,
      isBuiltIn,
      isOwn,
      registeredKey: key,
      pluginCount,
    });
    if (slug) coveredSlugs.add(slug);
  }

  // ── Step 2: cache-only dirs not covered by config ─────────────────────────
  if (fs.existsSync(cacheDir)) {
    const entries = fs.readdirSync(cacheDir, { withFileTypes: true });
    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      if (coveredSlugs.has(entry.name)) continue;

      const dirPath = path.join(cacheDir, entry.name);
      const manifest = findMarketplaceManifest(dirPath);
      const isBuiltIn = BUILTIN_SLUGS.some(s => entry.name.includes(s));
      const isOwn = entry.name === OCA_MARKETPLACE_SLUG;
      const pluginCount = Array.isArray(manifest?.plugins) ? manifest.plugins.length : 0;

      result.push({
        slug: entry.name,
        name: manifest?.name ?? entry.name,
        source: manifest?.source ?? entry.name,
        isBuiltIn,
        isOwn,
        registeredKey: null,
        pluginCount,
      });
    }
  }

  return result;
}

// ─── removeMarketplace ────────────────────────────────────────────────────────

/**
 * Remove a marketplace.
 *
 * - Registered (registeredKey non-null): unregister via CLI, then delete cache dir.
 * - Cache-only (registeredKey null): delete cache dir only.
 * - OCA marketplace and built-ins are always protected.
 *
 * @param {string} cacheDir        - path to ~/.copilot/marketplace-cache
 * @param {string} slug            - cache directory name
 * @param {string|null} registeredKey - config key, or null for cache-only
 * @returns {{ success: boolean, message: string }}
 */
export function removeMarketplace(cacheDir, slug, registeredKey) {
  if (slug === OCA_MARKETPLACE_SLUG || registeredKey === OCA_MARKETPLACE_KEY) {
    return { success: false, message: 'Cannot remove the office-coding-agent marketplace' };
  }
  if (BUILTIN_SLUGS.some(s => slug.includes(s)) || BUILTIN_KEYS.includes(registeredKey ?? '')) {
    return { success: false, message: 'Cannot remove built-in marketplaces' };
  }

  const errors = [];

  // Step 1: unregister from CLI if registered
  if (registeredKey) {
    try {
      execSync(`copilot plugin marketplace remove ${registeredKey}`, {
        encoding: 'utf-8',
        timeout: 30000,
        stdio: ['pipe', 'pipe', 'pipe'],
      });
    } catch (err) {
      errors.push(`CLI unregister failed: ${err.stderr?.trim() || err.message}`);
    }
  }

  // Step 2: delete cache directory
  const dirPath = path.join(cacheDir, slug);
  if (fs.existsSync(dirPath)) {
    try {
      fs.rmSync(dirPath, { recursive: true, force: true });
    } catch (err) {
      errors.push(`Cache delete failed: ${err.message}`);
    }
  }

  if (errors.length > 0) {
    return { success: false, message: errors.join('; ') };
  }
  return { success: true, message: `Removed marketplace: ${registeredKey ?? slug}` };
}
