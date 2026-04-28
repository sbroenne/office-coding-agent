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
  /** Skill names explicitly disabled by the user. Empty = all enabled. */
  disabledSkillNames: string[];
  /** MCP server names explicitly disabled by the user. Empty = all enabled. */
  disabledMcpServerNames: string[];
}

/** Default settings applied on first run */
export const DEFAULT_SETTINGS: UserSettings = {
  activeModel: 'claude-sonnet-4.6',
  disabledSkillNames: [],
  disabledMcpServerNames: [],
};
