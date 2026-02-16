import { createAzure } from '@ai-sdk/azure';
import { generateText } from 'ai';
import type { FoundryEndpoint, ModelInfo, ModelProvider } from '@/types';
import { normalizeEndpoint } from './aiClientFactory';

/** Cached discovery results per endpoint ID */
const modelCache = new Map<string, { result: DiscoveryResult; fetchedAt: number }>();

/** Cache TTL in milliseconds (5 minutes) */
const CACHE_TTL_MS = 5 * 60 * 1000;

/** Result of model discovery */
export interface DiscoveryResult {
  /** Discovered model list (always empty — manual entry required) */
  models: ModelInfo[];
  /** Always `'manual'` — AI Services has no data-plane discovery API */
  method: 'manual';
}

/**
 * "Discover" deployed models on a Foundry endpoint.
 *
 * AI Services / AI Foundry resources don't expose a data-plane deployments API,
 * so this always returns an empty result with `method: 'manual'`.
 * The UI prompts the user to enter model deployment names.
 *
 * Results are cached for 5 minutes per endpoint.
 */
// eslint-disable-next-line @typescript-eslint/require-await
export async function discoverModels(
  endpoint: FoundryEndpoint,
  forceRefresh = false
): Promise<DiscoveryResult> {
  if (!forceRefresh) {
    const cached = modelCache.get(endpoint.id);
    if (cached && Date.now() - cached.fetchedAt < CACHE_TTL_MS) {
      return cached.result;
    }
  }

  const result: DiscoveryResult = { models: [], method: 'manual' };
  modelCache.set(endpoint.id, { result, fetchedAt: Date.now() });
  return result;
}

/**
 * Validate that a model deployment name actually works by making
 * a minimal chat completion call via the AzureOpenAI SDK.
 *
 * Uses the SDK instead of raw `fetch` so it works in the Excel
 * WebView (which enforces CORS). Raw fetch would throw a CORS
 * error that was silently swallowed, making validation always fail.
 *
 * Returns `true` if the model responds, `false` otherwise.
 */
export async function validateModelDeployment(
  endpoint: FoundryEndpoint,
  modelId: string
): Promise<boolean> {
  const baseUrl = normalizeEndpoint(endpoint.resourceUrl);

  try {
    const provider = createAzure({
      baseURL: baseUrl + '/openai',
      apiKey: endpoint.apiKey ?? '',
    });

    await generateText({
      model: provider.chat(modelId),
      prompt: 'hi',
      maxOutputTokens: 5,
    });
    return true;
  } catch (err) {
    console.warn(`[validateModelDeployment] ${modelId} failed:`, err);
    return false;
  }
}

/** Clear the model cache for an endpoint */
export function clearModelCache(endpointId?: string): void {
  if (endpointId) {
    modelCache.delete(endpointId);
  } else {
    modelCache.clear();
  }
}

/** Check if a model ID looks like a non-chat model */
export function isEmbeddingOrUtilityModel(modelId: string): boolean {
  const lower = modelId.toLowerCase();
  return (
    lower.includes('embedding') ||
    lower.includes('dall-e') ||
    lower.includes('tts') ||
    lower.includes('whisper') ||
    lower.includes('transcribe') ||
    lower.includes('sora') ||
    lower.includes('gpt-image')
  );
}

/** Infer the provider from model ID and owned_by field */
export function inferProvider(modelId: string, ownedBy?: string): ModelProvider {
  const id = modelId.toLowerCase();
  const owner = (ownedBy ?? '').toLowerCase();

  if (id.includes('claude') || owner.includes('anthropic')) return 'Anthropic';
  if (id.includes('gpt-') || /\bo[134]-/.test(id) || /\bo[134]\b/.test(id)) return 'OpenAI';
  if (id.includes('deepseek') || id.includes('mai-ds')) return 'DeepSeek';
  if (id.includes('llama') || owner.includes('meta')) return 'Meta';
  if (id.includes('mistral') || id.includes('ministral')) return 'Mistral';
  if (id.includes('grok')) return 'xAI';
  if (id.includes('phi') || owner.includes('microsoft')) return 'Microsoft';
  return 'Other';
}

/** Format a model deployment ID into a human-readable name */
export function formatModelName(modelId: string): string {
  return modelId
    .replace(/-/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase())
    .replace(/Gpt /gi, 'GPT-')
    .replace(/Claude /gi, 'Claude ')
    .replace(/Deepseek /gi, 'DeepSeek ')
    .replace(/Mai Ds /gi, 'MAI-DS-');
}
