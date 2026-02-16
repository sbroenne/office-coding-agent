/**
 * Integration tests against a live Azure AI Foundry endpoint.
 *
 * These tests exercise the real HTTP API — no mocks.
 * They validate that our provider factory and chat
 * completions actually work end-to-end using the Vercel AI SDK.
 *
 * Configuration via environment variables:
 *   FOUNDRY_ENDPOINT  – Full resource URL (may include /api/projects/...)
 *   FOUNDRY_API_KEY   – API key for the endpoint
 *   FOUNDRY_MODEL     – Model deployment name to test (default: gpt-5.2-chat)
 *
 * Run:
 *   FOUNDRY_ENDPOINT=https://... FOUNDRY_API_KEY=... npx vitest run --config vitest.integration.config.ts
 *
 * These tests are excluded from the default test run (see vitest.config.ts).
 */

import { describe, it, expect, beforeAll } from 'vitest';
import { createAzure, type AzureOpenAIProvider } from '@ai-sdk/azure';
import { generateText, streamText, tool } from 'ai';
import { z } from 'zod';
import { normalizeEndpoint } from '@/services/ai/aiClientFactory';
import { createTestProvider, TEST_CONFIG } from '../test-provider';

// ─── Tests ───────────────────────────────────────────────────

describe.skipIf(!!TEST_CONFIG.skipReason)('Foundry Integration', { retry: 2 }, () => {
  let provider: AzureOpenAIProvider;

  beforeAll(async () => {
    provider = await createTestProvider();
  });

  // ── 1. Endpoint normalization ────────────────────────────

  describe('Endpoint normalization', () => {
    it('should strip project path from Foundry URL', () => {
      const normalized = normalizeEndpoint(
        'https://sbroenne.services.ai.azure.com/api/projects/proj-default'
      );
      expect(normalized).toBe('https://sbroenne.services.ai.azure.com');
    });

    it('should keep plain resource URLs unchanged', () => {
      const normalized = normalizeEndpoint('https://sbroenne.services.ai.azure.com');
      expect(normalized).toBe('https://sbroenne.services.ai.azure.com');
    });
  });

  // ── 2. Provider factory ──────────────────────────────────

  describe('Azure provider creation', () => {
    it('should create a provider that can generate text', async () => {
      const { text } = await generateText({
        model: provider.chat(TEST_CONFIG.model),
        prompt: 'Say hello',
        maxOutputTokens: 10,
      });

      expect(text.length).toBeGreaterThan(0);
      console.log('  Provider test response:', text.trim());
    });
  });

  // ── 4. Chat completions via AI SDK ───────────────────────

  describe('Chat completions', () => {
    it('should complete a simple chat request (non-streaming)', async () => {
      const { text } = await generateText({
        model: provider.chat(TEST_CONFIG.model),
        system: 'Respond with exactly one word.',
        messages: [{ role: 'user', content: 'Say hello.' }],
        maxOutputTokens: 100,
      });

      expect(text).toBeTruthy();
      expect(text.length).toBeGreaterThan(0);
      console.log('  Non-streaming response:', text.trim());
    });

    it('should stream chat completion chunks', async () => {
      const result = streamText({
        model: provider.chat(TEST_CONFIG.model),
        system: 'Respond with exactly one word.',
        messages: [{ role: 'user', content: 'Say hello.' }],
        maxOutputTokens: 100,
      });

      let content = '';
      let chunkCount = 0;

      for await (const part of result.fullStream) {
        if (part.type === 'text-delta') {
          chunkCount++;
          content += part.text;
        }
      }

      expect(chunkCount).toBeGreaterThan(0);
      expect(content.length).toBeGreaterThan(0);
      console.log(`  Streamed ${chunkCount} chunks, content: "${content.trim()}"`);
    });

    it('should handle tool calling (non-streaming)', async () => {
      const weatherTool = tool({
        description: 'Get the current weather for a location',
        inputSchema: z.object({
          location: z.string().describe('City name'),
        }),
        execute: async ({ location }) => ({
          location,
          temperature: 72,
          condition: 'sunny',
        }),
      });

      const { text, toolCalls } = await generateText({
        model: provider.chat(TEST_CONFIG.model),
        messages: [{ role: 'user', content: 'What is the weather in Seattle?' }],
        tools: { get_weather: weatherTool },
        maxOutputTokens: 200,
      });

      // Model should have called the tool
      if (toolCalls.length > 0) {
        expect(toolCalls[0].toolName).toBe('get_weather');
        expect(toolCalls[0].input).toHaveProperty('location');
        console.log('  Tool call:', toolCalls[0].toolName, toolCalls[0].input);
      } else {
        expect(text.length).toBeGreaterThan(0);
        console.log('  Model responded without tool call:', text.trim());
      }
    });

    it('should handle tool calling (streaming)', async () => {
      const weatherTool = tool({
        description: 'Get the current weather for a location',
        inputSchema: z.object({
          location: z.string().describe('City name'),
        }),
        execute: async ({ location }) => ({
          location,
          temperature: 72,
          condition: 'sunny',
        }),
      });

      const result = streamText({
        model: provider.chat(TEST_CONFIG.model),
        messages: [{ role: 'user', content: 'What is the weather in London?' }],
        tools: { get_weather: weatherTool },
        maxOutputTokens: 200,
      });

      let content = '';
      const toolCallNames: string[] = [];

      for await (const part of result.fullStream) {
        if (part.type === 'text-delta') {
          content += part.text;
        } else if (part.type === 'tool-call') {
          toolCallNames.push(part.toolName);
        }
      }

      if (toolCallNames.length > 0) {
        expect(toolCallNames[0]).toBe('get_weather');
        console.log('  Streamed tool call:', toolCallNames[0]);
      } else {
        expect(content.length).toBeGreaterThan(0);
        console.log('  Streamed content (no tool call):', content.trim());
      }
    });
  });

  // ── 5. Error handling ────────────────────────────────────

  // Tests that create providers with known-bad API keys to verify error paths.
  // These require us to know the base URL, which works regardless of auth method.

  describe('Error handling', () => {
    it('should reject with a clear error for invalid API key', async () => {
      const baseUrl = normalizeEndpoint(TEST_CONFIG.endpoint);
      const badProvider = createAzure({
        baseURL: baseUrl + '/openai',
        apiKey: 'invalid-key-12345',
      });

      await expect(
        generateText({
          model: badProvider.chat(TEST_CONFIG.model),
          prompt: 'test',
          maxOutputTokens: 5,
        })
      ).rejects.toThrow();
    });

    it('should reject for a non-existent model deployment', async () => {
      await expect(
        generateText({
          model: provider.chat('nonexistent-model-xyz-12345'),
          prompt: 'test',
          maxOutputTokens: 5,
        })
      ).rejects.toThrow();
    });
  });
});
