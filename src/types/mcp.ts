/** Transport type for an MCP server */
export type McpTransportType = 'http' | 'sse' | 'stdio';

/** Lifecycle status of an MCP server */
export type McpServerStatus =
  | 'stopped'
  | 'starting'
  | 'connected'
  | 'needs-auth'
  | 'pending'
  | 'disabled'
  | 'not_configured'
  | 'failed'
  | 'error';

/** A single log entry from an MCP server */
export interface McpLogEntry {
  timestamp: string;
  level: 'info' | 'warn' | 'error';
  message: string;
}

/** Runtime state tracked per MCP server (ephemeral, not persisted) */
export interface McpServerState {
  status: McpServerStatus;
  error?: string;
  tools: { name: string; description: string }[];
  logs: McpLogEntry[];
}

/** A configured MCP server imported from a mcp.json file */
export interface McpServerConfig {
  /** Display name (used as identifier) */
  name: string;
  /** Optional description shown in the UI */
  description?: string;
  /** Transport protocol */
  transport: McpTransportType;
  /** MCP server endpoint URL (required for http/sse transport) */
  url?: string;
  /** Optional HTTP headers (e.g. Authorization) — for http/sse transport */
  headers?: Record<string, string>;
  /** Executable command (required for stdio transport, e.g. "npx") */
  command?: string;
  /** Command arguments (for stdio transport) */
  args?: string[];
  /** Optional environment variables (for stdio transport) */
  env?: Record<string, string>;
}
