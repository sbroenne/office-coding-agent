/** Transport type for an MCP server (browser-compatible options only; no stdio) */
export type McpTransportType = 'http' | 'sse';

/** A configured MCP server imported from a mcp.json file */
export interface McpServerConfig {
  /** Display name (used as identifier) */
  name: string;
  /** Optional description shown in the UI */
  description?: string;
  /** MCP server endpoint URL */
  url: string;
  /** Transport protocol */
  transport: McpTransportType;
  /** Optional HTTP headers (e.g. Authorization) */
  headers?: Record<string, string>;
}
