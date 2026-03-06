/**
 * marketplaceService.mjs
 *
 * Testable helper functions for marketplace listing and removal.
 * Extracted from server.mjs so the logic can be unit-tested without Express.
 */

import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';

// ─── Constants ────────────────────────────────────────────────────────────────

export const OCA_MARKETPLACE_KEY = 'office-coding-agent';
export const BUILTIN_KEYS = ['copilot-plugins', 'awesome-copilot'];

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
 * Convert an "owner/repo" string to all plausible cache-dir name forms that
 * the Copilot CLI might use:
 *   "owner/repo"         → ["owner-repo", "owner--repo"]
 *   "stbrnner_ms/SPT-IQ" → ["stbrnner_ms-SPT-IQ", "stbrnner_ms--SPT-IQ",
 *                            "stbrnner-ms-spt-iq"]
 * We try multiple forms because the CLI's exact algorithm is not public.
 */
export function repoCacheSlugs(repo) {
  const slugs = new Set();
  // Form 1: replace first "/" with "-" (original case)
  slugs.add(repo.replace('/', '-'));
  // Form 2: replace first "/" with "--" (original case, observed for repos with _ in owner)
  slugs.add(repo.replace('/', '--'));
  // Form 3: full lowercase slugify (replace all non-alphanumeric with "-")
  slugs.add(repo.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, ''));
  // Form 4: keep owner chars, slugify only repo
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
 * List all explicitly registered marketplaces from config.json.
 * Cache-only dirs (never registered) are intentionally ignored.
 *
 * @param {string} cacheDir   - path to ~/.copilot/marketplace-cache
 * @param {string} configPath - path to ~/.copilot/config.json
 * @returns {{ slug, name, source, isBuiltIn, isOwn, registeredKey, pluginCount }[]}
 */
export function listMarketplaces(cacheDir, configPath) {
  const config = readConfig(configPath);
  const configMarketplaces = config.marketplaces || {};
  const result = [];

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
  }

  return result;
}

/**
 * Remove a registered marketplace via the Copilot CLI.
 *
 * @param {string} registeredKey - config key (e.g. "my-marketplace")
 * @returns {{ success: boolean, message: string }}
 */
export function removeMarketplace(registeredKey) {
  if (registeredKey === OCA_MARKETPLACE_KEY) {
    return { success: false, message: 'Cannot remove the office-coding-agent marketplace' };
  }
  if (BUILTIN_KEYS.includes(registeredKey)) {
    return { success: false, message: 'Cannot remove built-in marketplaces' };
  }

  try {
    execSync(`copilot plugin marketplace remove ${registeredKey}`, {
      encoding: 'utf-8',
      timeout: 30000,
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    return { success: true, message: `Removed marketplace: ${registeredKey}` };
  } catch (err) {
    return { success: false, message: err.stderr?.trim() || err.message };
  }
}
