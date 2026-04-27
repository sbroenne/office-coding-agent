import { describe, expect, it } from 'vitest';

import { toMcpServerKey } from '@/utils/mcpServerKey';

describe('toMcpServerKey', () => {
  it.each([
    ['powerbi', 'powerbi'],
    ['Power BI MCP', 'power_bi_mcp'],
    [' powerbi ', 'powerbi'],
    ['Power  BI', 'power_bi'],
  ])('normalizes %s to %s', (name, expected) => {
    expect(toMcpServerKey(name)).toBe(expected);
  });
});
