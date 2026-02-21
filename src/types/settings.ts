import type { AgentConfig } from './agent';
import type { AgentSkill } from './skill';
import type { McpServerConfig } from './mcp';

/** Provider labels for grouping models in the picker */
export type ModelProvider = 'Anthropic' | 'OpenAI' | 'Google' | 'Other';

/** A Copilot-supported model option */
export interface CopilotModel {
  id: string;
  name: string;
  provider: ModelProvider;
}

/** Infer provider from model ID prefix */
export function inferProvider(modelId: string): ModelProvider {
  if (modelId.startsWith('claude')) return 'Anthropic';
  if (
    modelId.startsWith('gpt') ||
    modelId.startsWith('o1') ||
    modelId.startsWith('o3') ||
    modelId.startsWith('o4')
  )
    return 'OpenAI';
  if (modelId.startsWith('gemini')) return 'Google';
  return 'Other';
}

/** Persisted user settings */
export interface UserSettings {
  /** Currently selected Copilot model ID */
  activeModel: string;
  /** Names of currently active agent skills. null = all skills enabled (default). */
  activeSkillNames: string[] | null;
  /** ID of the currently selected agent (matches agent metadata name). */
  activeAgentId: string;
  /** Imported skills loaded from local ZIP files. */
  importedSkills: AgentSkill[];
  /** Imported agents loaded from local ZIP files. */
  importedAgents: AgentConfig[];
  /** MCP servers imported from a mcp.json file. */
  importedMcpServers: McpServerConfig[];
  /** Names of currently active MCP servers. null = all servers enabled (default). */
  activeMcpServerNames: string[] | null;
}

/** Default settings applied on first run */
export const DEFAULT_SETTINGS: UserSettings = {
  activeModel: 'claude-sonnet-4.6',
  activeSkillNames: null,
  activeAgentId: 'Excel',
  importedSkills: [],
  importedAgents: [],
  importedMcpServers: [],
  activeMcpServerNames: null,
};
