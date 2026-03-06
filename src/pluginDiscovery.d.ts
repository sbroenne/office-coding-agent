/**
 * Type declarations for src/pluginDiscovery.mjs (plain JS module).
 */
declare module '*/pluginDiscovery.mjs' {
  export function slugify(name: string): string;
  export function isPluginForHost(pluginName: string, host: string): boolean;
  export function readCopilotConfig(configPath?: string): Promise<Record<string, unknown>>;
  export function readPluginManifest(pluginDir: string): Promise<Record<string, unknown> | null>;
  export function discoverPluginSkillDirs(host?: string, configPath?: string): Promise<string[]>;
  export function discoverPluginAgents(
    host?: string,
    configPath?: string
  ): Promise<Array<{ name: string; description: string; prompt: string }>>;
  export const COPILOT_CONFIG_PATH: string;
}
