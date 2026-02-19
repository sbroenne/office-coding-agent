/**
 * General-purpose tools available across all Office hosts.
 *
 * These tools are not tied to any specific Office host and are injected
 * into every agent alongside the host-specific tools.
 */

import { tool, generateText, stepCountIs, type LanguageModel, type ToolSet } from 'ai';
import { z } from 'zod';

/** Default maximum response length for web_fetch */
const DEFAULT_MAX_LENGTH = 10_000;

/** Fetch the content of a URL and return it as text. */
export const webFetchTool = tool({
  description:
    'Fetch the content of a URL and return it as text. Use this to retrieve web pages, JSON APIs, or any publicly accessible HTTP resource.',
  inputSchema: z.object({
    url: z.string().describe('The URL to fetch'),
    maxLength: z
      .number()
      .optional()
      .describe(
        `Maximum number of characters to return. Defaults to ${String(DEFAULT_MAX_LENGTH)}.`
      ),
  }),
  execute: async ({ url, maxLength = DEFAULT_MAX_LENGTH }) => {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`HTTP ${String(response.status)}: ${response.statusText}`);
    }
    const text = await response.text();
    return text.length > maxLength ? text.slice(0, maxLength) + '… [truncated]' : text;
  },
});

/** Default maximum tool-call steps for run_subagent */
const DEFAULT_SUBAGENT_MAX_STEPS = 5;

/**
 * Create a `run_subagent` tool bound to a specific model and tool set.
 *
 * The subagent receives the host tools (e.g. Excel commands) but intentionally
 * does NOT receive `run_subagent` itself to prevent unbounded recursion.
 */
export function createRunSubagentTool(model: LanguageModel, hostTools: ToolSet) {
  return tool({
    description:
      "Delegate a focused subtask to an AI subagent that has access to the same host tools (e.g. Excel commands). Returns the subagent's text response. Use this to break complex multi-step work into isolated sub-tasks.",
    inputSchema: z.object({
      task: z.string().describe('The task or question for the subagent to complete'),
      systemPrompt: z
        .string()
        .optional()
        .describe("Optional system prompt to guide the subagent's behaviour"),
      maxSteps: z
        .number()
        .int()
        .min(1)
        .optional()
        .describe(
          `Maximum number of tool-call steps the subagent may take. Defaults to ${String(DEFAULT_SUBAGENT_MAX_STEPS)}.`
        ),
    }),
    execute: async ({ task, systemPrompt, maxSteps = DEFAULT_SUBAGENT_MAX_STEPS }) => {
      const result = await generateText({
        model,
        ...(systemPrompt ? { system: systemPrompt } : {}),
        prompt: task,
        tools: hostTools,
        stopWhen: stepCountIs(maxSteps),
      });
      return result.text;
    },
  });
}

/**
 * Build the general-purpose tool set for a given model and host.
 *
 * The returned tools are host-agnostic and safe to merge into any agent's
 * tool set. The `run_subagent` tool is given the host tools plus `web_fetch`
 * so it can fetch web content, but it intentionally does NOT receive
 * `run_subagent` itself to prevent unbounded recursive delegation.
 */
export function getGeneralTools(model: LanguageModel, hostTools: ToolSet): ToolSet {
  // Tools the subagent can use: host tools + web_fetch, but NOT run_subagent
  const subagentTools: ToolSet = { ...hostTools, web_fetch: webFetchTool };
  return {
    web_fetch: webFetchTool,
    run_subagent: createRunSubagentTool(model, subagentTools),
  };
}
