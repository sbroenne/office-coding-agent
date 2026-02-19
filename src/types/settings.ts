/** Authentication method for an endpoint */
import type { AgentConfig } from './agent';
import type { AgentSkill } from './skill';
import type { McpServerConfig } from './mcp';

/** Authentication method for an endpoint */
export type AuthMethod = 'apiKey';

/** Provider backend for an endpoint */
export type ProviderType = 'azure' | 'anthropic' | 'openai' | 'mistral' | 'deepseek' | 'xai';

/** Represents a configured AI provider endpoint */
export interface FoundryEndpoint {
  /** Unique identifier */
  id: string;
  /** User-friendly display name (e.g., "East US Production") */
  displayName: string;
  /** Base resource URL (e.g., "https://my-resource.openai.azure.com") */
  resourceUrl: string;
  /** How this endpoint authenticates */
  authMethod: AuthMethod;
  /** API key — required when authMethod is 'apiKey' */
  apiKey?: string;
  /** Provider backend. Defaults to 'azure' when not set (backward-compatible). */
  providerType?: ProviderType;
}

/** Information about a deployed model discovered from the endpoint */
export interface ModelInfo {
  /** Model deployment name / ID used in API calls */
  id: string;
  /** Human-readable model name */
  name: string;
  /** Owner/provider (e.g., "system", "organization") */
  ownedBy: string;
  /** Inferred provider label for grouping in UI */
  provider: ModelProvider;
  /** Whether this is the user's configured default */
  isDefault?: boolean;
}

/** Known model providers for UI grouping */
export type ModelProvider =
  | 'Anthropic'
  | 'OpenAI'
  | 'DeepSeek'
  | 'Meta'
  | 'Mistral'
  | 'xAI'
  | 'Microsoft'
  | 'Other';

/** Persisted user settings */
export interface UserSettings {
  /** Configured Foundry endpoints */
  endpoints: FoundryEndpoint[];
  /** Currently active endpoint ID */
  activeEndpointId: string | null;
  /** Currently selected model deployment name */
  activeModelId: string | null;
  /** Default model to select on startup */
  defaultModelId: string;
  /** Per-endpoint model lists (endpointId → ModelInfo[]) */
  endpointModels: Record<string, ModelInfo[]>;
  /** Names of currently active agent skills. null = all skills enabled (default). */
  activeSkillNames: string[] | null;
  /** ID of the currently selected agent (matches agent metadata name). */
  activeAgentId: string;
  /** Imported skills loaded from local ZIP files. Bundled skills are managed separately and read-only. */
  importedSkills: AgentSkill[];
  /** Imported agents loaded from local ZIP files. Bundled agents are managed separately and read-only. */
  importedAgents: AgentConfig[];
  /** MCP servers imported from a mcp.json file. */
  importedMcpServers: McpServerConfig[];
  /** Names of currently active MCP servers. null = all servers enabled (default). */
  activeMcpServerNames: string[] | null;
}

/** Default settings applied on first run */
export const DEFAULT_SETTINGS: UserSettings = {
  endpoints: [],
  activeEndpointId: null,
  activeModelId: null,
  defaultModelId: 'gpt-5.2-chat',
  endpointModels: {},
  activeSkillNames: null,
  activeAgentId: 'Excel',
  importedSkills: [],
  importedAgents: [],
  importedMcpServers: [],
  activeMcpServerNames: null,
};
