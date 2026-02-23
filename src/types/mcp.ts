/** Transport type for an MCP server */
export type McpTransportType = 'http' | 'sse' | 'stdio';

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
