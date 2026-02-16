/** A suggested prompt shown to the user */
export interface Suggestion {
  /** Unique suggestion ID */
  id: string;
  /** Display text */
  text: string;
  /** Optional icon */
  icon?: string;
}

/** A chat message in the conversation history */
export interface ChatMessage {
  /** Message role */
  role: 'user' | 'assistant' | 'tool';
  /** Text content */
  content: string;
  /** Whether this message is currently being streamed */
  isStreaming?: boolean;
  /** Tool calls made by the assistant in this message */
  toolCalls?: ToolCall[];
  /** For tool-result messages: the ID of the tool call this result belongs to */
  toolCallId?: string;
}

/** A tool call made by the assistant */
export interface ToolCall {
  /** Unique tool call ID */
  id: string;
  /** Name of the function that was called */
  functionName: string;
  /** Serialized JSON arguments */
  arguments: string;
  /** Parsed argument object */
  parsedArguments?: Record<string, unknown>;
}
