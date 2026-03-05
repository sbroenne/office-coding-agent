/** Status of a tool call within a message */
export interface ToolCallStatus {
  type: 'running' | 'complete' | 'incomplete';
  reason?: 'cancelled' | 'error';
  error?: unknown;
}

/** A tool-call part within a chat message */
export interface ToolCallPart {
  type: 'tool-call';
  toolCallId: string;
  toolName: string;
  argsText: string;
  result?: string;
  status?: ToolCallStatus;
}

/** A text part within a chat message */
export interface TextPart {
  type: 'text';
  text: string;
}

/** Union of all message content parts */
export type ChatMessagePart = TextPart | ToolCallPart;

/** Overall status of a chat message */
export type ChatMessageStatus =
  | { type: 'running' }
  | { type: 'complete'; reason: 'stop' }
  | { type: 'incomplete'; reason: 'cancelled' | 'error'; error?: string };

/** A single message in the chat thread */
export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: ChatMessagePart[];
  status?: ChatMessageStatus;
  /** Per-message thinking/working text (shimmer label). Null = not thinking. */
  thinkingText?: string | null;
  createdAt: Date;
}
