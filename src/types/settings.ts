import type { AgentConfig } from './agent';
import type { AgentSkill } from './skill';
import type { McpServerConfig } from './mcp';

/** A Copilot-supported model option */
export interface CopilotModel {
  id: string;
  name: string;
  provider: 'Anthropic' | 'OpenAI' | 'Google' | 'Other';
}

/** Available Copilot models */
export const COPILOT_MODELS: CopilotModel[] = [
  { id: 'claude-sonnet-4.5', name: 'Claude Sonnet 4.5', provider: 'Anthropic' },
  { id: 'claude-opus-4.5', name: 'Claude Opus 4.5', provider: 'Anthropic' },
  { id: 'gpt-4.1', name: 'GPT-4.1', provider: 'OpenAI' },
  { id: 'gpt-4o', name: 'GPT-4o', provider: 'OpenAI' },
  { id: 'o3-mini', name: 'o3-mini', provider: 'OpenAI' },
  { id: 'gemini-2.0-flash', name: 'Gemini 2.0 Flash', provider: 'Google' },
];

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
  activeModel: 'claude-sonnet-4.5',
  activeSkillNames: null,
  activeAgentId: 'Excel',
  importedSkills: [],
  importedAgents: [],
  importedMcpServers: [],
  activeMcpServerNames: null,
};
