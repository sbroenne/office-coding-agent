/**
 * Type declarations for src/pluginDiscovery.mjs (plain JS module).
 */
declare module '*/pluginDiscovery.mjs' {
  export class ConfigReadError extends Error {
    cause: unknown;
  }
  export function slugify(name: string): string;
  export function isPluginForHost(pluginName: string, host: string): boolean;
  export function parsePluginAgentFrontmatter(raw: string): {
    description: string;
    hosts: string[];
    tools?: string[];
  };
  export function readCopilotConfig(configPath?: string): Promise<Record<string, unknown>>;
  export function listAllPluginConfigs(configPath?: string): Promise<Array<Record<string, unknown>>>;
  export function readPluginManifest(pluginDir: string): Promise<Record<string, unknown> | null>;
  export function discoverPluginSkillDirs(host?: string, configPath?: string): Promise<string[]>;
  export function discoverPluginAgents(
    host?: string,
    configPath?: string
  ): Promise<Array<{ name: string; description: string; prompt: string; hosts: string[]; tools?: string[] }>>;
  export function discoverPluginSkillObjects(
    host?: string,
    configPath?: string
  ): Promise<Array<{ name: string; description: string; version: string; hosts: string[]; content: string }>>;
  export function discoverPluginPrompts(
    host?: string,
    configPath?: string
  ): Promise<Array<{ name: string; description: string; agent: string; argumentHint: string; body: string }>>;
  export function readMcpServersForPlugin(
    plugin: Record<string, unknown>,
    manifest?: Record<string, unknown> | null
  ): Promise<Record<string, unknown>>;
  export const COPILOT_CONFIG_PATH: string;
}
