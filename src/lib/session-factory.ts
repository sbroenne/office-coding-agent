/**
 * Factory for creating focused Copilot sub-sessions.
 *
 * Used by the deck orchestrator to spawn planner and worker sessions
 * on the same WebSocket connection as the main chat session.
 */

import type { WebSocketCopilotClient, BrowserCopilotSession } from '@/lib/websocket-client';
import type { Tool } from '@github/copilot-sdk';
import type { SessionEvent } from '@github/copilot-sdk';

export interface SubSessionConfig {
  /** Model ID (e.g. "gpt-4.1") */
  model: string;
  /** Full system prompt for this sub-session */
  systemPrompt: string;
  /** Tools available to this sub-session (empty for planner) */
  tools?: Tool[];
}

export interface SubSessionResult {
  /** All text content from the assistant */
  text: string;
  /** All events received during the session */
  events: SessionEvent[];
  /** Whether the session completed successfully */
  success: boolean;
  /** Error message if failed */
  error?: string;
}

/**
 * Create a sub-session, send a prompt, collect all events, and destroy it.
 *
 * @param client - The shared WebSocketCopilotClient
 * @param config - Session configuration (model, prompt, tools)
 * @param prompt - The user prompt to send
 * @param onEvent - Optional callback for real-time event streaming
 * @returns All collected text and events
 */
export async function runSubSession(
  client: WebSocketCopilotClient,
  config: SubSessionConfig,
  prompt: string,
  onEvent?: (event: SessionEvent) => void,
): Promise<SubSessionResult> {
  let session: BrowserCopilotSession | null = null;
  const events: SessionEvent[] = [];
  let text = '';

  try {
    session = await client.createSession({
      model: config.model,
      systemMessage: { mode: 'replace', content: config.systemPrompt },
      tools: config.tools ?? [],
    });

    for await (const event of session.query({ prompt })) {
      events.push(event);
      onEvent?.(event);

      if (event.type === 'assistant.message_delta') {
        text += event.data.deltaContent;
      } else if (event.type === 'assistant.message') {
        text = event.data.content;
      } else if (event.type === 'session.error') {
        return {
          text,
          events,
          success: false,
          error: event.data.message,
        };
      }
    }

    return { text, events, success: true };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return { text, events, success: false, error: msg };
  } finally {
    if (session) {
      try {
        await session.destroy();
      } catch {
        // ignore cleanup errors
      }
    }
  }
}
