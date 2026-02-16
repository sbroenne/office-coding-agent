export {
  getAzureProvider,
  invalidateClient,
  clearAllClients,
  normalizeEndpoint,
} from './aiClientFactory';
export { sendChatMessage, messagesToCoreMessages } from './chatService';
export type { ChatRequestOptions } from './chatService';
export {
  discoverModels,
  clearModelCache,
  validateModelDeployment,
  inferProvider,
  isEmbeddingOrUtilityModel,
  formatModelName,
} from './modelDiscoveryService';
export type { DiscoveryResult } from './modelDiscoveryService';
export { BASE_PROMPT, getAppPromptForHost, buildSystemPrompt } from './systemPrompt';
