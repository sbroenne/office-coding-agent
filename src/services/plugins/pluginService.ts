/**
 * Browser-side service for the Plugin Hub.
 * Wraps /api/plugins/* REST endpoints with typed responses.
 */

import type {
  InstalledPlugin,
  RegisteredMarketplace,
  BrowsePlugin,
  PluginManifest,
  PluginComponents,
  PluginActionResult,
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

function postJson<T>(url: string, body: Record<string, string>): Promise<T> {
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
export async function getMarketplaces(): Promise<RegisteredMarketplace[]> {
  const data = await fetchJson<{ marketplaces: RegisteredMarketplace[] }>(
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
export async function uninstallPlugin(name: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/uninstall`, { name });
}

/** Enable a disabled plugin. */
export async function enablePlugin(name: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/enable`, { name });
}

/** Disable a plugin without uninstalling it. */
export async function disablePlugin(name: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/disable`, { name });
}

/** Update a plugin to the latest version. */
export async function updatePlugin(name: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/update`, { name });
}

/** Register a new marketplace. */
export async function addMarketplace(spec: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/marketplace/add`, { spec });
}

/** Remove a registered marketplace. */
export async function removeMarketplace(registeredKey: string): Promise<PluginActionResult> {
  return postJson<PluginActionResult>(`${API_BASE}/marketplace/remove`, { registeredKey });
}
