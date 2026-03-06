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
  /** ID of the currently selected agent (matches agent metadata name). */
  activeAgentId: string;
}

/** Default settings applied on first run */
export const DEFAULT_SETTINGS: UserSettings = {
  activeModel: 'claude-sonnet-4.6',
  activeAgentId: 'Excel',
};

/** Built-in MCP servers that ship with the add-in. Non-removable, but toggleable. */
export const BUNDLED_MCP_SERVERS: McpServerConfig[] = [
  {
    name: 'workiq',
    description: 'Microsoft 365 Copilot — emails, meetings, documents, Teams',
    transport: 'stdio',
    command: 'npx',
    args: ['-y', '@microsoft/workiq', 'mcp'],
  },
  {
    name: 'powerbi',
    description: 'Power BI — query semantic models, generate DAX, explore data',
    transport: 'http',
    url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
  },
];
