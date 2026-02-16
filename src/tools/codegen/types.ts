/**
 * Declarative tool configuration types.
 *
 * Each ToolConfig defines everything needed to generate:
 *   1. A Vercel AI SDK `tool()` with Zod inputSchema
 *   2. An Excel command function (excelRun + getSheet + load + sync)
 *   3. A manifest entry for pytest-aitest MCP testing
 *
 * The config IS the single source of truth — no hand-written tool or command files.
 */

// ─── Parameter Types ──────────────────────────────────────

/** Supported Zod types for tool parameters */
export type ParamType = 'string' | 'number' | 'boolean' | 'string[]' | 'any[][]' | 'string[][]';

/** A single tool parameter definition */
export interface ParamDef {
  /** Zod type */
  type: ParamType;
  /** Whether the parameter is required (default: true) */
  required?: boolean;
  /** LLM-facing description */
  description: string;
  /** For string enums — allowed values */
  enum?: readonly string[];
  /** Default value (makes the param optional at runtime) */
  default?: unknown;
}

// ─── Tool Config ──────────────────────────────────────────

/**
 * Complete declarative definition of one Excel tool.
 *
 * For simple patterns (get/set/create/delete/list/action), the factory
 * generates the full command implementation from `execute`.
 *
 * The `execute` function receives the Excel RequestContext and typed args,
 * and returns the result data. It runs inside `excelRun()` automatically.
 */
export interface ToolConfig {
  /** Tool name as the LLM sees it (e.g., "get_range_values") */
  name: string;

  /** LLM-facing description — what this tool does and when to use it */
  description: string;

  /** Parameter definitions → generates Zod inputSchema */
  params: Record<string, ParamDef>;

  /**
   * The command implementation.
   * Receives (context, args) — runs inside excelRun() automatically.
   * Return the result data (not ToolCallResult — the factory wraps it).
   */
  execute: (context: Excel.RequestContext, args: Record<string, unknown>) => Promise<unknown>;
}

// ─── Manifest Types (for pytest-aitest) ───────────────────

/** JSON-serializable param definition for the manifest */
export interface ManifestParam {
  type: ParamType;
  required: boolean;
  description: string;
  enum?: readonly string[];
  default?: unknown;
}

/** JSON-serializable tool definition for the manifest */
export interface ManifestTool {
  name: string;
  description: string;
  params: Record<string, ManifestParam>;
}

/** The full manifest exported as JSON */
export interface ToolManifest {
  version: string;
  generatedAt: string;
  tools: ManifestTool[];
}
