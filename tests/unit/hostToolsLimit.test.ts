import { describe, expect, it } from 'vitest';
import { getToolsForHost, MAX_TOOLS_PER_REQUEST } from '@/tools';

describe('host tool limits', () => {
  it('caps excel tools at the provider maximum', () => {
    const tools = getToolsForHost('excel');
    expect(Object.keys(tools).length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('returns empty tools for unknown host', () => {
    const tools = getToolsForHost('unknown');
    expect(Object.keys(tools)).toHaveLength(0);
  });
});
