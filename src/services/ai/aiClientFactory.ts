import { createAzure } from '@ai-sdk/azure';
import type { AzureOpenAIProvider } from '@ai-sdk/azure';
import type { FoundryEndpoint } from '@/types';

/** Cache of Azure provider instances per endpoint ID */
const providerCache = new Map<string, AzureOpenAIProvider>();

/**
 * Create or retrieve a cached Azure OpenAI provider for a Foundry endpoint.
 *
 * The provider is cached per endpoint ID — the API key doesn't change between requests.
 *
 * @param endpoint - The Foundry endpoint configuration
 * @returns An AI SDK Azure provider instance
 * @throws {Error} If the endpoint configuration is invalid
 */
export function getAzureProvider(endpoint: FoundryEndpoint): AzureOpenAIProvider {
  const cached = providerCache.get(endpoint.id);
  if (cached) return cached;

  // Validate endpoint configuration
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

  providerCache.set(endpoint.id, provider);
  return provider;
}

/** Invalidate a cached provider (e.g., when endpoint config changes) */
export function invalidateClient(endpointId: string): void {
  providerCache.delete(endpointId);
}

/** Clear all cached providers */
export function clearAllClients(): void {
  providerCache.clear();
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
