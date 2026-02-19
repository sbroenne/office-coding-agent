import type { ProviderType } from '@/types';

/** Metadata for a provider — used by the setup wizard and settings dialog */
export interface ProviderConfig {
  /** Provider type identifier */
  type: ProviderType;
  /** Display label shown in the UI */
  label: string;
  /**
   * Fixed API base URL.
   * Undefined for Azure — the user must provide their own resource URL.
   */
  baseUrl?: string;
  /** Well-known model IDs for this provider, pre-populated in the wizard */
  defaultModels: string[];
}

export const PROVIDER_CONFIGS: Record<ProviderType, ProviderConfig> = {
  azure: {
    type: 'azure',
    label: 'Azure AI Foundry',
    defaultModels: [],
  },
  openai: {
    type: 'openai',
    label: 'OpenAI',
    baseUrl: 'https://api.openai.com/v1',
    defaultModels: ['gpt-4o', 'gpt-4o-mini', 'o3', 'o4-mini'],
  },
  anthropic: {
    type: 'anthropic',
    label: 'Anthropic',
    baseUrl: 'https://api.anthropic.com',
    defaultModels: ['claude-opus-4-5', 'claude-sonnet-4-5', 'claude-haiku-3-5'],
  },
  mistral: {
    type: 'mistral',
    label: 'Mistral AI',
    baseUrl: 'https://api.mistral.ai',
    defaultModels: ['mistral-large-latest', 'mistral-small-latest', 'codestral-latest'],
  },
  deepseek: {
    type: 'deepseek',
    label: 'DeepSeek',
    baseUrl: 'https://api.deepseek.com',
    defaultModels: ['deepseek-chat', 'deepseek-reasoner'],
  },
  xai: {
    type: 'xai',
    label: 'xAI',
    baseUrl: 'https://api.x.ai',
    defaultModels: ['grok-3-beta', 'grok-2-1212'],
  },
};

/** Ordered list of provider types for display in the UI */
export const PROVIDER_ORDER: ProviderType[] = [
  'azure',
  'openai',
  'anthropic',
  'mistral',
  'deepseek',
  'xai',
];

/** Type-safe accessor for provider config — avoids ESLint no-unsafe-member-access on computed keys */
export function getProviderConfig(type: ProviderType): ProviderConfig {
  return PROVIDER_CONFIGS[type];
}
