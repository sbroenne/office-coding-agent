/**
 * TypeScript types for the Copilot CLI plugin system.
 *
 * These map directly to the data structures in:
 * - ~/.copilot/config.json (installed_plugins[])
 * - ~/.copilot/installed-plugins/<marketplace>/<name>/plugin.json
 * - ~/.copilot/marketplace-cache/<slug>/.github/plugin/marketplace.json
 */

/**
 * A parsed prompt template from a plugin's prompts/*.prompt.md file.
 * Slash commands in the ChatComposer are built from these.
 */
export interface PluginPrompt {
  /** Slash command name (filename without .prompt.md) */
  name: string;
  /** One-line description shown in the slash menu */
  description: string;
  /** The agent to activate when this prompt is selected (from "agent:" frontmatter) */
  agent: string;
  /** Placeholder text for the variable, e.g. "TPID (e.g. 12345678)" */
  argumentHint: string;
  /** Markdown body — injected into composer input on selection */
  body: string;
}

/** Author metadata in plugin.json and marketplace.json */
export interface PluginAuthor {
  name: string;
  email?: string;
  url?: string;
}

/**
 * Plugin manifest — maps to `plugin.json` fields.
 * See: https://docs.github.com/en/copilot/reference/cli-plugin-reference#pluginjson
 */
export interface PluginManifest {
  name: string;
  description?: string;
  version?: string;
  author?: PluginAuthor;
  homepage?: string;
  repository?: string;
  license?: string;
  keywords?: string[];
  category?: string;
  tags?: string[];

  // Component path fields (all optional, CLI uses defaults if omitted)
  agents?: string | string[];
  skills?: string | string[];
  commands?: string | string[];
  hooks?: string | object;
  mcpServers?: string | object;
  lspServers?: string | object;
}

/**
 * Installed plugin entry from ~/.copilot/config.json → installed_plugins[].
 */
export interface InstalledPluginConfig {
  name: string;
  marketplace: string; // empty string for direct installs
  version: string;
  installed_at: string; // ISO 8601
  enabled: boolean;
  cache_path: string; // absolute path to plugin directory
}

/**
 * Enriched installed plugin — config entry + resolved manifest + component counts.
 * This is what the Plugin Hub UI renders.
 */
export interface InstalledPlugin {
  name: string;
  marketplace: string;
  version: string;
  enabled: boolean;
  installedAt: string;
  cachePath: string;
  manifest: PluginManifest | null;
  components: PluginComponents;
}

/** Counts and names of components discovered in a plugin's directory. */
export interface PluginComponents {
  agentCount: number;
  agentNames: string[];
  skillCount: number;
  skillNames: string[];
  mcpServerCount: number;
  mcpServerNames: string[];
  hookCount: number;
  commandCount: number;
}

/**
 * Plugin source in marketplace.json — can be a relative path string,
 * or an object with source type + repo + path for external repos.
 */
export type PluginSource = string | { source: string; repo?: string; path?: string };

/**
 * Marketplace plugin entry — from marketplace.json → plugins[].
 * See: https://docs.github.com/en/copilot/reference/cli-plugin-reference#marketplacejson
 */
export interface MarketplacePlugin {
  name: string;
  description?: string;
  version?: string;
  source: PluginSource;
  author?: PluginAuthor;
  homepage?: string;
  repository?: string;
  license?: string;
  keywords?: string[];
  category?: string;
  tags?: string[];
  agents?: string | string[];
  skills?: string | string[];
  commands?: string | string[];
  hooks?: string | object;
  mcpServers?: string | object;
  lspServers?: string | object;
  strict?: boolean;
}

/** Marketplace metadata block in marketplace.json */
export interface MarketplaceMetadata {
  description?: string;
  version?: string;
  pluginRoot?: string;
}

/** Marketplace owner block in marketplace.json */
export interface MarketplaceOwner {
  name: string;
  email?: string;
}

/**
 * Full marketplace manifest — maps to marketplace.json.
 */
export interface PluginMarketplace {
  name: string;
  owner: MarketplaceOwner;
  metadata?: MarketplaceMetadata;
  plugins: MarketplacePlugin[];
}

/**
 * Registered marketplace info (from CLI or config).
 */
export interface RegisteredMarketplace {
  slug: string; // cache directory name, e.g. "sbroenne-office-coding-agent-plugins"
  name: string;
  source: string; // e.g. "github/copilot-plugins" or local path
  isBuiltIn: boolean; // true for copilot-plugins and awesome-copilot
  pluginCount?: number;
  /** The key in ~/.copilot/config.json marketplaces — required for removal. Null if not registered. */
  registeredKey: string | null;
  /** True if this is the office-coding-agent's own marketplace (cannot be removed). */
  isOwn: boolean;
}

/**
 * Plugin install specification — how to reference a plugin for installation.
 */
export type PluginInstallSpec =
  | { type: 'marketplace'; plugin: string; marketplace: string } // plugin@marketplace
  | { type: 'github'; owner: string; repo: string; path?: string } // owner/repo[:path]
  | { type: 'url'; url: string } // https://...
  | { type: 'local'; path: string }; // ./path or /abs/path

/**
 * Result of a plugin mutation (install/uninstall/enable/disable/update).
 */
export interface PluginActionResult {
  success: boolean;
  message: string;
  plugin?: string;
}

/**
 * Browse result — marketplace plugin with install status.
 */
export interface BrowsePlugin extends MarketplacePlugin {
  marketplace: string;
  installed: boolean;
  enabled?: boolean;
}
