import { createAzure } from '@ai-sdk/azure';
import type { AzureOpenAIProvider } from '@ai-sdk/azure';
import { createAnthropic } from '@ai-sdk/anthropic';
import { createOpenAI } from '@ai-sdk/openai';
import { createMistral } from '@ai-sdk/mistral';
import { createDeepSeek } from '@ai-sdk/deepseek';
import { createXai } from '@ai-sdk/xai';
import type { LanguageModel } from 'ai';
import type { FoundryEndpoint } from '@/types';

/** Cache of provider instances per endpoint ID */
const azureProviderCache = new Map<string, AzureOpenAIProvider>();
const genericProviderCache = new Map<string, ReturnType<typeof createOpenAI>>();

/** Anthropic API base URL */
export const ANTHROPIC_BASE_URL = 'https://api.anthropic.com';

/**
 * Create or retrieve a cached Azure OpenAI provider for a Foundry endpoint.
 *
 * @param endpoint - The Foundry endpoint configuration
 * @returns An AI SDK Azure provider instance
 * @throws {Error} If the endpoint configuration is invalid
 */
export function getAzureProvider(endpoint: FoundryEndpoint): AzureOpenAIProvider {
  const cached = azureProviderCache.get(endpoint.id);
  if (cached) return cached;

  if (!endpoint.resourceUrl || endpoint.resourceUrl.trim() === '') {
    throw new Error('Endpoint URL is required. Please configure it in Settings.');
  }
  if (!endpoint.apiKey || endpoint.apiKey.trim() === '') {
    throw new Error('API Key is required. Please configure it in Settings.');
  }

  const normalizedUrl = normalizeEndpoint(endpoint.resourceUrl);
  console.log('[aiClientFactory] Creating Azure provider:', {
    endpointId: endpoint.id,
    baseURL: normalizedUrl + '/openai',
    hasApiKey: !!endpoint.apiKey,
  });

  const provider = createAzure({
    baseURL: normalizedUrl + '/openai',
    apiKey: endpoint.apiKey,
  });

  azureProviderCache.set(endpoint.id, provider);
  return provider;
}

/** @deprecated Use getProviderModel() instead */
export const getAnthropicProvider = (endpoint: FoundryEndpoint) =>
  createAnthropic({ apiKey: endpoint.apiKey ?? '' });

/**
 * Get a language model for the given endpoint and model ID.
 * Routes to the appropriate provider based on `endpoint.providerType`.
 *
 * @param endpoint - The configured endpoint
 * @param modelId  - The model ID / deployment name
 * @returns A LanguageModel ready for use with the AI SDK
 */
export function getProviderModel(endpoint: FoundryEndpoint, modelId: string): LanguageModel {
  const type = endpoint.providerType ?? 'azure';
  const apiKey = endpoint.apiKey ?? '';

  switch (type) {
    case 'anthropic':
      return createAnthropic({ apiKey })(modelId);
    case 'openai':
      return createOpenAI({ apiKey })(modelId);
    case 'mistral':
      return createMistral({ apiKey })(modelId);
    case 'deepseek':
      return createDeepSeek({ apiKey })(modelId);
    case 'xai':
      return createXai({ apiKey })(modelId);
    case 'azure':
    default:
      return getAzureProvider(endpoint).chat(modelId);
  }
}

/** Invalidate a cached provider (e.g., when endpoint config changes) */
export function invalidateClient(endpointId: string): void {
  azureProviderCache.delete(endpointId);
  genericProviderCache.delete(endpointId);
}

/** Clear all cached providers */
export function clearAllClients(): void {
  azureProviderCache.clear();
  genericProviderCache.clear();
}

/** Normalize the endpoint URL — strip project paths and suffixes to get the base resource URL */
export function normalizeEndpoint(resourceUrl: string): string {
  let url = resourceUrl.trim();
  // Remove trailing slashes
  while (url.endsWith('/')) url = url.slice(0, -1);
  // Remove /openai/v1 suffix if user pasted a full URL
  url = url.replace(/\/openai\/v1\/?$/, '');
  // Remove /openai suffix to avoid doubling
  url = url.replace(/\/openai\/?$/, '');
  // Remove Foundry project path (e.g., /api/projects/proj-default)
  url = url.replace(/\/api\/projects\/[^/]+$/, '');
  return url;
}
