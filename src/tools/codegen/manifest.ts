/**
 * Manifest exporter — generates a JSON manifest from ToolConfig arrays.
 *
 * The manifest is the bridge between the TypeScript add-in and the
 * Python pytest-aitest MCP server. It carries tool names, descriptions,
 * and parameter schemas so the Python side can dynamically register
 * identical tools backed by an in-memory spreadsheet simulator.
 */

import type { ToolConfig, ToolManifest, ManifestTool, ManifestParam } from './types';

/** Convert a ToolConfig to a JSON-serializable ManifestTool */
function toManifestTool(config: ToolConfig): ManifestTool {
  const params: Record<string, ManifestParam> = {};

  for (const [key, def] of Object.entries(config.params)) {
    params[key] = {
      type: def.type,
      required: def.required !== false,
      description: def.description,
      ...(def.enum && { enum: def.enum }),
      ...(def.default !== undefined && { default: def.default }),
    };
  }

  return {
    name: config.name,
    description: config.description,
    params,
  };
}

/**
 * Generate a ToolManifest from multiple config arrays.
 * Called by the manifest generation script at build time.
 */
export function generateManifest(...configArrays: (readonly ToolConfig[])[]): ToolManifest {
  const tools: ManifestTool[] = [];

  for (const configs of configArrays) {
    for (const config of configs) {
      tools.push(toManifestTool(config));
    }
  }

  return {
    version: '1.0.0',
    generatedAt: new Date().toISOString(),
    tools,
  };
}
