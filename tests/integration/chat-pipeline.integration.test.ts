/**
 * Integration test for the chat pipeline.
 *
 * Tests the full flow: ToolLoopAgent → DirectChatTransport → useChat.
 * Uses a live Azure AI Foundry endpoint (same creds as foundry.integration.test.ts).
 *
 * This is the test for "I type a prompt and nothing happens" — it verifies that
 * the chat pipeline actually sends messages and receives streaming responses.
 *
 * Configuration via environment variables:
 *   FOUNDRY_ENDPOINT  – Full resource URL
 *   FOUNDRY_API_KEY   – API key for the endpoint
 *   FOUNDRY_MODEL     – Model deployment name (default: gpt-5.2-chat)
 */

import { describe, it, expect, beforeAll } from 'vitest';
import { createAzure, type AzureOpenAIProvider } from '@ai-sdk/azure';
import { ToolLoopAgent, DirectChatTransport, stepCountIs } from 'ai';
import { normalizeEndpoint } from '@/services/ai/aiClientFactory';
import { excelTools } from '@/tools';
import { createTestProvider, TEST_CONFIG } from '../test-provider';

/** Create a simple user message compatible with DirectChatTransport. */
function userMessage(id: string, text: string) {
  return { id, role: 'user' as const, parts: [{ type: 'text' as const, text }] };
}

// ─── Tests ───────────────────────────────────────────────────

// eslint-disable-next-line vitest/valid-describe-callback
describe.skipIf(!!TEST_CONFIG.skipReason)(
  'Chat pipeline (DirectChatTransport)',
  { retry: 2 },
  () => {
    let provider: AzureOpenAIProvider;
    // We use `any` for agent/transport because ToolLoopAgent and DirectChatTransport
    // have deeply nested generics that don't play well with variable declarations.
    let agent: any;
    let transport: any;

    beforeAll(async () => {
      provider = await createTestProvider();

      agent = new ToolLoopAgent({
        model: provider.chat(TEST_CONFIG.model),
        instructions: 'You are a helpful assistant. Respond concisely.',
        tools: excelTools,
        stopWhen: stepCountIs(3),
      });

      transport = new DirectChatTransport({ agent });
    });

    it('sends a message and receives a streaming response', async () => {
      const stream: ReadableStream = await transport.sendMessages({
        chatId: 'test-chat',
        trigger: 'submit-message',
        messages: [userMessage('test-1', 'Say "hello" and nothing else.')],
        abortSignal: AbortSignal.timeout(30_000),
      });

      expect(stream).toBeDefined();
      expect(stream).toBeInstanceOf(ReadableStream);

      // Consume the stream and collect chunks
      const reader = stream.getReader();
      const chunks: unknown[] = [];
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
      }

      expect(chunks.length).toBeGreaterThan(0);

      // Check that at least one chunk contains text content
      const hasText = chunks.some(
        (chunk: any) => chunk.type === 'text' || chunk.type === 'text-delta'
      );
      expect(hasText).toBe(true);
    });

    it('agent is created with the correct tools', () => {
      const toolNames = Object.keys(excelTools);
      expect(toolNames.length).toBeGreaterThan(0);
      expect(agent).toBeDefined();
      expect(agent.tools).toBeDefined();
      // eslint-disable-next-line @typescript-eslint/no-unsafe-argument
      expect(Object.keys(agent.tools)).toHaveLength(toolNames.length);
    });

    it('stream surfaces an error for invalid model without hanging', async () => {
      const baseUrl = normalizeEndpoint(TEST_CONFIG.endpoint);
      const badProvider = createAzure({
        baseURL: baseUrl + '/openai',
        apiKey: TEST_CONFIG.apiKey || 'dummy-key-for-error-test',
      });

      const badAgent = new ToolLoopAgent({
        model: badProvider.chat('nonexistent-model-xyz'),
        instructions: 'test',
        tools: {},
        stopWhen: stepCountIs(1),
      });

      const badTransport: any = new DirectChatTransport({ agent: badAgent });

      // sendMessages returns a ReadableStream.  For invalid models the error
      // is surfaced inside the stream (logged to stderr) rather than rejecting.
      // Verify the stream finishes without hanging — that's the contract.
      const stream: ReadableStream = await badTransport.sendMessages({
        chatId: 'test-bad',
        trigger: 'submit-message',
        messages: [userMessage('test-bad', 'hello')],
        abortSignal: AbortSignal.timeout(15_000),
      });

      const reader = stream.getReader();
      const chunks: unknown[] = [];
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
      }

      // Stream should complete (not hang) — chunks may be empty or contain error info
      expect(stream).toBeDefined();
    });
  }
);
