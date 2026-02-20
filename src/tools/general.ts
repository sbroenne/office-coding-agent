/**
 * General-purpose tools available across all Office hosts.
 *
 * These tools are not tied to any specific Office host and are injected
 * into every agent alongside the host-specific tools.
 */

import type { Tool, ToolInvocation, ToolResultObject } from '@github/copilot-sdk';

/** Default maximum response length for web_fetch */
const DEFAULT_MAX_LENGTH = 10_000;

/** Fetch the content of a URL and return it as text. */
export const webFetchTool: Tool = {
  name: 'web_fetch',
  description:
    'Fetch the content of a URL and return it as text. Use this to retrieve web pages, JSON APIs, or any publicly accessible HTTP resource.',
  parameters: {
    type: 'object',
    properties: {
      url: { type: 'string', description: 'The URL to fetch' },
      maxLength: {
        type: 'number',
        description: `Maximum number of characters to return. Defaults to ${String(DEFAULT_MAX_LENGTH)}.`,
      },
    },
    required: ['url'],
  },
  handler: async (
    args: unknown,
    _invocation: ToolInvocation
  ): Promise<ToolResultObject | string> => {
    const proxyUrl = `/api/fetch?url=${encodeURIComponent((args as { url: string }).url)}`;
    const response = await fetch(proxyUrl);
    if (!response.ok) throw new Error(`HTTP ${String(response.status)}: ${response.statusText}`);
    const text = await response.text();
    const maxLength = (args as { maxLength?: number }).maxLength ?? DEFAULT_MAX_LENGTH;
    return text.length > maxLength ? text.slice(0, maxLength) + '… [truncated]' : text;
  },
};
