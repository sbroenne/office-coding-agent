/**
 * Tool factory — generates Vercel AI SDK tools from declarative ToolConfig.
 *
 * Each config entry produces a `tool()` with:
 *   - A Zod inputSchema built from the `params` definition
 *   - An execute fn that wraps config.execute() inside excelRun()
 */

import { tool, type Tool } from 'ai';
import { z, type ZodTypeAny } from 'zod';
import type { ToolConfig, ParamType } from './types';
import { excelRun, getSheet } from '@/services/excel/helpers';

// Re-export getSheet so configs can use it without extra imports
export { getSheet };

/** Map ParamType → Zod schema */
function zodForType(type: ParamType): ZodTypeAny {
  switch (type) {
    case 'string':
      return z.string();
    case 'number':
      return z.number();
    case 'boolean':
      return z.boolean();
    case 'string[]':
      return z.array(z.string());
    case 'any[][]':
      return z.array(z.array(z.any()));
    case 'string[][]':
      return z.array(z.array(z.string()));
  }
}

/** Build a Zod object schema from a ParamDef record */
function buildZodSchema(params: ToolConfig['params']): ZodTypeAny {
  const shape: Record<string, ZodTypeAny> = {};

  for (const [key, def] of Object.entries(params)) {
    let schema = zodForType(def.type);

    // Apply enum constraint for strings
    if (def.enum && def.type === 'string') {
      schema = z.enum(def.enum as [string, ...string[]]);
    }

    // Add description
    schema = schema.describe(def.description);

    // Make optional if not required
    if (def.required === false) {
      schema = schema.optional();
    }

    shape[key] = schema;
  }

  return z.object(shape);
}

/**
 * Create Vercel AI SDK tools from an array of ToolConfig.
 * Returns a Record<toolName, Tool> ready to spread into excelTools.
 */
export function createTools(configs: readonly ToolConfig[]): Record<string, Tool> {
  const tools: Record<string, Tool> = {};

  for (const config of configs) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-assignment -- dynamic schema can't satisfy tool()'s generic inference
    const schema = buildZodSchema(config.params) as any;
    tools[config.name] = tool({
      description: config.description,
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment -- schema is dynamically built
      inputSchema: schema,
      execute: async (args: Record<string, unknown>) => {
        return excelRun(async context => {
          return config.execute(context, args);
        });
      },
    });
  }

  return tools;
}
