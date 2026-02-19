/**
 * Tool factory — generates Copilot SDK tools from declarative ToolConfig.
 *
 * Each config entry produces a Tool with:
 *   - A JSON Schema parameters object built from the `params` definition
 *   - A handler that wraps config.execute() inside excelRun()
 */

import type { Tool, ToolInvocation, ToolResultObject } from '@github/copilot-sdk';
import type { ToolConfig, ParamType } from './types';
import { excelRun, getSheet } from '@/services/excel/helpers';

// Re-export getSheet so configs can use it without extra imports
export { getSheet };

type JSONSchemaProperty = Record<string, unknown>;

/** Map ParamType → JSON Schema property object */
function jsonSchemaForType(type: ParamType, enumValues?: readonly string[]): JSONSchemaProperty {
  switch (type) {
    case 'string':
      return enumValues ? { type: 'string', enum: enumValues } : { type: 'string' };
    case 'number':
      return { type: 'number' };
    case 'boolean':
      return { type: 'boolean' };
    case 'string[]':
      return { type: 'array', items: { type: 'string' } };
    case 'any[][]':
      return { type: 'array', items: { type: 'array' } };
    case 'string[][]':
      return { type: 'array', items: { type: 'array', items: { type: 'string' } } };
  }
}

/** Build a JSON Schema object from a ParamDef record */
function buildJsonSchema(params: ToolConfig['params']): Record<string, unknown> {
  const properties: Record<string, JSONSchemaProperty> = {};
  const required: string[] = [];

  for (const [key, def] of Object.entries(params)) {
    properties[key] = {
      ...jsonSchemaForType(def.type, def.enum),
      description: def.description,
    };

    if (def.required !== false && def.default === undefined) {
      required.push(key);
    }
  }

  return {
    type: 'object',
    properties,
    ...(required.length > 0 ? { required } : {}),
  };
}

/**
 * Create Copilot SDK tools from an array of ToolConfig.
 * Returns Tool[] ready for use in SessionConfig.tools.
 */
export function createTools(configs: readonly ToolConfig[]): Tool[] {
  return configs.map(config => ({
    name: config.name,
    description: config.description,
    parameters: buildJsonSchema(config.params),
    handler: async (_args: unknown, invocation: ToolInvocation): Promise<ToolResultObject> => {
      const args = invocation.arguments as Record<string, unknown>;
      try {
        const result = await excelRun(context => config.execute(context, args));
        return {
          textResultForLlm: typeof result === 'string' ? result : JSON.stringify(result),
          resultType: 'success',
          toolTelemetry: {},
        };
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return {
          textResultForLlm: message,
          resultType: 'failure',
          error: message,
          toolTelemetry: {},
        };
      }
    },
  }));
}
