import {
  streamText,
  stepCountIs,
  type LanguageModel,
  type ModelMessage,
  type ToolCallPart,
  type ToolResultPart,
  type ToolSet,
} from 'ai';
import type { ChatMessage, ToolCall, ToolCallResult } from '@/types';

import { getToolsForHost } from '@/tools';
import { buildSkillContext } from '@/services/skills';
import { detectOfficeHost, type OfficeHostApp } from '@/services/office/host';
import { buildSystemPrompt } from './systemPrompt';

/** Options for a chat request */
export interface ChatRequestOptions {
  /** Model deployment name */
  modelId: string;
  /** Full message history */
  messages: ChatMessage[];
  /** Callback for streaming content chunks */
  onContent?: (chunk: string) => void;
  /** Callback when tool calls are detected */
  onToolCalls?: (toolCalls: ToolCall[]) => void;
  /** Callback when tool results are available */
  onToolResult?: (toolCallId: string, result: ToolCallResult) => void;
  /** Callback when streaming completes */
  onComplete?: (fullContent: string) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
  /** Abort signal */
  signal?: AbortSignal;
  /** Office host application for host-specific tool/prompt routing */
  host?: OfficeHostApp;
  /**
   * Optional tool set — defaults to host-selected tools.
   * Pass simulated tools in integration tests to exercise the full
   * sendChatMessage pipeline without a live Office host instance.
   */
  tools?: ToolSet;
}

/**
 * Send a chat completion request with streaming and tool-calling support.
 * Uses AI SDK's streamText with automatic multi-step tool execution.
 */
export async function sendChatMessage(
  model: LanguageModel,
  options: ChatRequestOptions
): Promise<string> {
  const { messages, onContent, onToolCalls, onToolResult, onComplete, onError, signal } = options;

  try {
    const coreMessages = messagesToCoreMessages(messages);
    const skillContext = buildSkillContext();
    const host = options.host ?? detectOfficeHost();

    const result = streamText({
      model,
      system: buildSystemPrompt(host) + skillContext,
      messages: coreMessages,
      tools: options.tools ?? getToolsForHost(host),
      stopWhen: stepCountIs(10),
      abortSignal: signal,
    });

    let fullContent = '';

    for await (const part of result.fullStream) {
      switch (part.type) {
        case 'text-delta':
          fullContent += part.text;
          onContent?.(part.text);
          break;

        case 'tool-call': {
          const toolCall: ToolCall = {
            id: part.toolCallId,
            functionName: part.toolName,
            arguments: JSON.stringify(part.input),
            parsedArguments: part.input as Record<string, unknown>,
          };
          onToolCalls?.([toolCall]);
          break;
        }

        case 'tool-result': {
          const toolResult = part.output as ToolCallResult;
          onToolResult?.(part.toolCallId, toolResult);
          break;
        }

        case 'error':
          throw part.error;
      }
    }

    onComplete?.(fullContent);
    return fullContent;
  } catch (error) {
    const err = error instanceof Error ? error : new Error(String(error));
    onError?.(err);
    throw err;
  }
}

/** Convert our ChatMessage[] to AI SDK CoreMessage[] format */
export function messagesToCoreMessages(messages: ChatMessage[]): ModelMessage[] {
  const result: ModelMessage[] = [];

  for (const m of messages) {
    if (m.isStreaming) continue;

    if (m.role === 'user') {
      result.push({ role: 'user', content: m.content });
    } else if (m.role === 'assistant') {
      if (m.toolCalls && m.toolCalls.length > 0) {
        const toolCallParts: ToolCallPart[] = m.toolCalls.map(tc => ({
          type: 'tool-call' as const,
          toolCallId: tc.id,
          toolName: tc.functionName,
          input: (tc.parsedArguments ?? JSON.parse(tc.arguments)) as Record<string, unknown>,
        }));
        const content: (ToolCallPart | { type: 'text'; text: string })[] = [];
        if (m.content) {
          content.push({ type: 'text', text: m.content });
        }
        content.push(...toolCallParts);
        result.push({ role: 'assistant', content });
      } else {
        result.push({ role: 'assistant', content: m.content });
      }
    } else if (m.role === 'tool' && m.toolCallId) {
      // Find if there's already a tool message we can merge into
      const lastMsg = result[result.length - 1];
      const toolResultPart: ToolResultPart = {
        type: 'tool-result',
        toolCallId: m.toolCallId,
        toolName: '', // AI SDK doesn't require this for history
        output: {
          type: 'text' as const,
          value: typeof m.content === 'string' ? m.content : JSON.stringify(m.content),
        },
      };
      if (lastMsg?.role === 'tool') {
        (lastMsg.content as ToolResultPart[]).push(toolResultPart);
      } else {
        result.push({ role: 'tool', content: [toolResultPart] });
      }
    }
  }

  return result;
}
