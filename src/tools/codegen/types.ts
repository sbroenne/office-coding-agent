/**
 * Declarative tool configuration types.
 *
 * Each ToolConfig defines everything needed to generate:
 *   1. A Copilot SDK `Tool` with JSON Schema parameters
 *   2. An Office command function (context.run + load + sync)
 *   3. A manifest entry for pytest-aitest MCP testing
 *
 * The config IS the single source of truth — no hand-written tool or command files.
 * TContext controls which Office host RequestContext is used.
 */

// ─── Parameter Types ──────────────────────────────────────

/** Supported types for tool parameters */
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
 * Structural definition shared by all host-specific tool configs.
 * Used by manifest generation (which doesn't need the execute function).
 */
export interface ToolConfigBase {
  /** Tool name as the LLM sees it (e.g., "get_range_values") */
  name: string;

  /** LLM-facing description — what this tool does and when to use it */
  description: string;

  /** Parameter definitions → generates JSON Schema */
  params: Record<string, ParamDef>;
}

/**
 * Complete declarative definition of one Office tool.
 *
 * TContext is the Office RequestContext type for the target host:
 *   - Excel (default): ToolConfig or ToolConfig<Excel.RequestContext>
 *   - PowerPoint:      ToolConfig<PowerPoint.RequestContext>  (PptToolConfig)
 *   - Word:            ToolConfig<Word.RequestContext>         (WordToolConfig)
 *
 * The `execute` function receives the host RequestContext and typed args,
 * and returns the result data. It is wrapped inside the appropriate
 * `Office.run()` call automatically by the factory.
 */
export interface ToolConfig<TContext = Excel.RequestContext> extends ToolConfigBase {
  /**
   * The command implementation.
   * Receives (context, args) — runs inside Office.run() automatically.
   * Return the result data (not ToolCallResult — the factory wraps it).
   */
  execute: (context: TContext, args: Record<string, unknown>) => Promise<unknown>;
}

/** Convenience alias for PowerPoint tools */
export type PptToolConfig = ToolConfig<PowerPoint.RequestContext>;

/** Convenience alias for Word tools */
export type WordToolConfig = ToolConfig<Word.RequestContext>;

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
