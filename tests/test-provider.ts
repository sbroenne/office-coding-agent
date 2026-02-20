/**
 * Shared test provider factory for integration tests.
 *
 * Creates an Azure OpenAI provider using either:
 *   1. API key auth (FOUNDRY_API_KEY env var), or
 *   2. Entra ID / Azure CLI auth (@azure/identity DefaultAzureCredential)
 *
 * This lets developers run integration tests with just `az login` —
 * no API key needed.
 *
 * Usage:
 *   import { createTestProvider, TEST_CONFIG } from '../test-provider';
 *
 *   const provider = await createTestProvider();
 *   // provider.chat(TEST_CONFIG.model) → ready to use
 */

import { createAzure, type AzureOpenAIProvider } from '@ai-sdk/azure';
import { normalizeEndpoint } from '@/services/ai/aiClientFactory';

// ─── Configuration ───────────────────────────────────────────

const ENDPOINT = process.env.FOUNDRY_ENDPOINT ?? '';
const API_KEY = process.env.FOUNDRY_API_KEY ?? '';
const MODEL = process.env.FOUNDRY_MODEL ?? 'gpt-5.2-chat';

/** Exported config for tests to read */
export const TEST_CONFIG = {
  endpoint: ENDPOINT,
  apiKey: API_KEY,
  model: MODEL,
  /** True when an API key is available (some tests need it directly) */
  hasApiKey: !!API_KEY,
  /** True if we have enough config to run integration tests */
  get canRun(): boolean {
    return !!ENDPOINT;
  },
  /** Reason to skip if we can't run */
  get skipReason(): string | undefined {
    if (!ENDPOINT) return 'Set FOUNDRY_ENDPOINT env var to run integration tests';
    return undefined;
  },
} as const;

/**
 * Create an Azure OpenAI provider for integration tests.
 *
 * Prefers API key when available (faster, no token refresh).
 * Falls back to Entra ID via DefaultAzureCredential (az login, managed identity, etc.).
 */
export async function createTestProvider(): Promise<AzureOpenAIProvider> {
  if (!ENDPOINT) {
    throw new Error('FOUNDRY_ENDPOINT is required. Set it in .env');
  }

  const baseUrl = normalizeEndpoint(ENDPOINT);

  // Fast path: API key auth
  if (API_KEY) {
    return createAzure({
      baseURL: baseUrl + '/openai',
      apiKey: API_KEY,
    });
  }

  // Entra ID path: get a token from DefaultAzureCredential (az login, managed identity, etc.)
  const { DefaultAzureCredential } = await import('@azure/identity');
  const credential = new DefaultAzureCredential();

  const tokenResponse = await credential.getToken('https://cognitiveservices.azure.com/.default');
  if (!tokenResponse?.token) {
    throw new Error('Failed to acquire Entra ID token. Run `az login` or set FOUNDRY_API_KEY.');
  }

  // Azure OpenAI supports both `api-key` header and `Authorization: Bearer` header.
  // We pass an empty apiKey so createAzure doesn't set a real api-key header,
  // and provide the Bearer token via custom headers.
  return createAzure({
    baseURL: baseUrl + '/openai',
    apiKey: '', // empty — auth comes from Bearer header below
    headers: {
      Authorization: `Bearer ${tokenResponse.token}`,
    },
  });
}
