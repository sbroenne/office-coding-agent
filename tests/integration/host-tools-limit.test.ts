import { describe, it, expect } from 'vitest';
import { getToolsForHost, MAX_TOOLS_PER_REQUEST } from '@/tools';

describe('host tools limit', () => {
  it('getToolsForHost("excel") returns at most MAX_TOOLS_PER_REQUEST tools', () => {
    const tools = getToolsForHost('excel');
    expect(tools.length).toBeGreaterThan(0);
    expect(tools.length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('getToolsForHost("powerpoint") returns at most MAX_TOOLS_PER_REQUEST tools', () => {
    const tools = getToolsForHost('powerpoint');
    expect(tools.length).toBeGreaterThan(0);
    expect(tools.length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('getToolsForHost("word") returns at most MAX_TOOLS_PER_REQUEST tools', () => {
    const tools = getToolsForHost('word');
    expect(tools.length).toBeGreaterThan(0);
    expect(tools.length).toBeLessThanOrEqual(MAX_TOOLS_PER_REQUEST);
  });

  it('getToolsForHost("unknown") returns an empty array', () => {
    const tools = getToolsForHost('unknown' as never);
    expect(tools).toEqual([]);
  });
});
