/**
 * Browser-side service for the Plugin Hub.
 * Wraps /api/plugins/* REST endpoints with typed responses.
 */

import type {
  InstalledPlugin,
  BrowsePlugin,
  PluginManifest,
  PluginComponents,
  PluginActionResult,
  PluginMarketplaceSummary,
} from '@/types/plugin';
import { getLocalApiBase } from '@/lib/api';

const API_BASE = `${getLocalApiBase()}/api/plugins`;

async function fetchJson<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, init);
  if (!res.ok) {
    const text = await res.text().catch(() => res.statusText);
    throw new Error(`Plugin API error (${res.status}): ${text}`);
  }
  return res.json() as Promise<T>;
}

function postJson<T>(url: string, body: Record<string, string | null | undefined>): Promise<T> {
  return fetchJson<T>(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
}

/** Get all installed plugins with manifest data and component counts. */
export async function getInstalledPlugins(): Promise<InstalledPlugin[]> {
  const data = await fetchJson<{ plugins: InstalledPlugin[] }>(`${API_BASE}/installed`);
  return data.plugins;
}

/** Get all registered marketplaces. */
export async function getMarketplaces(): Promise<PluginMarketplaceSummary[]> {
  const data = await fetchJson<{ marketplaces: PluginMarketplaceSummary[] }>(
    `${API_BASE}/marketplaces`
  );
  return data.marketplaces;
}

/** Browse plugins in a marketplace. Returns plugins with install status. */
export async function browseMarketplace(marketplace: string): Promise<BrowsePlugin[]> {
  const data = await fetchJson<{ marketplace: string; plugins: BrowsePlugin[] }>(
    `${API_BASE}/browse/${encodeURIComponent(marketplace)}`
  );
  return data.plugins;
}

/** Get full details for an installed plugin. */
export async function getPluginDetails(name: string): Promise<{
  plugin: InstalledPlugin;
  manifest: PluginManifest | null;
  components: PluginComponents;
}> {
  // API returns flat {...plugin, manifest, components} — restructure for the UI
  const data = await fetchJson<
    InstalledPlugin & { manifest: PluginManifest | null; components: PluginComponents }
  >(`${API_BASE}/${encodeURIComponent(name)}/details`);
  const { manifest, components, ...plugin } = data;
  return { plugin: plugin as InstalledPlugin, manifest, components };
}

/** Install a plugin from a marketplace, repo, or local path. */
export async function installPlugin(spec: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/install`, { spec });
}

/** Uninstall a plugin. */
export async function uninstallPlugin(
  name: string,
  marketplace?: string
): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/uninstall`, { name, marketplace });
}

/** Update a plugin to the latest version. */
export async function updatePlugin(
  name: string,
  marketplace?: string
): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/update`, { name, marketplace });
}

/** Update all installed plugins. */
export async function updateAllPlugins(): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/update-all`, {});
}

/** Register a new marketplace. */
export async function addMarketplace(spec: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/marketplace/add`, { spec });
}

/** Refresh one marketplace or all marketplaces. */
export async function updateMarketplace(name?: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/marketplace/update`, { name });
}

/** Remove a marketplace — registered (by CLI) or cache-only (by slug). */
export async function removeMarketplace(
  slug: string,
  registeredKey?: string | null
): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/marketplace/remove`, {
    slug,
    registeredKey: registeredKey ?? null,
  });
}
