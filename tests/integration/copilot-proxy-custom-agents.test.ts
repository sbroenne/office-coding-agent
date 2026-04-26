// @vitest-environment node
import { describe, it, expect } from 'vitest';
import {
  applySessionToolAccessToCustomAgents,
  getRegisteredToolNames,
  mergePluginMcpServers,
} from '@/../src/copilotProxy.mjs';

describe('copilotProxy custom agent tool registration', () => {
  it('uses active session tool names when a custom agent omits tools', () => {
    const toolNames = getRegisteredToolNames(
      [{ name: 'range' }, { name: 'workbook' }, { name: 'manage_memory' }],
      undefined
    );

    const agents = applySessionToolAccessToCustomAgents(
      [{ name: 'Plugin Agent', prompt: 'Use the sheet tools.' }],
      toolNames
    );

    expect(agents[0].tools).toEqual(['range', 'workbook', 'manage_memory']);
  });

  it('expands wildcard or null tool access to the active session tools', () => {
    const toolNames = ['range', 'workbook'];

    const agents = applySessionToolAccessToCustomAgents(
      [
        { name: 'Wildcard Agent', prompt: 'All tools', tools: ['*'] },
        { name: 'Null Agent', prompt: 'All tools', tools: null },
      ],
      toolNames
    );

    expect(agents[0].tools).toEqual(toolNames);
    expect(agents[1].tools).toEqual(toolNames);
  });

  it('preserves explicit tool scoping against the active session tools', () => {
    const toolNames = getRegisteredToolNames(
      [{ name: 'range' }, { name: 'workbook' }, { name: 'manage_memory' }],
      ['range', 'manage_memory']
    );

    const agents = applySessionToolAccessToCustomAgents(
      [
        {
          name: 'Scoped Agent',
          prompt: 'Limited tools only',
          tools: ['range', 'workbook', 'missing_tool'],
        },
      ],
      toolNames
    );

    expect(toolNames).toEqual(['range', 'manage_memory']);
    expect(agents[0].tools).toEqual(['range']);
  });
});

describe('copilotProxy plugin MCP merge', () => {
  it('keeps built-in or user MCP servers when plugin servers use the same name', () => {
    const merged = mergePluginMcpServers(
      { workiq: { type: 'stdio', command: 'npx' } },
      { workiq: { type: 'stdio', command: 'plugin-workiq' }, pluginOnly: { type: 'http', url: 'https://example.test' } }
    );

    expect(merged.workiq).toEqual({ type: 'stdio', command: 'npx' });
    expect(merged.pluginOnly).toEqual({ type: 'http', url: 'https://example.test' });
  });
});
