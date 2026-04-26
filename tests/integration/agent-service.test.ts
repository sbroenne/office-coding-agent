import { describe, it, expect } from 'vitest';
import {
  parseAgentFrontmatter,
  getAgents,
  getPickerAgents,
  getAllAgents,
  getAgent,
  getAgentInstructions,
  getDefaultAgent,
  resolveActiveAgent,
} from '@/services/agents/agentService';

describe('agentService — parseAgentFrontmatter', () => {
  it('parses valid frontmatter', () => {
    const md = `---
name: Test Agent
description: A test agent
version: 1.0.0
hosts: [excel]
defaultForHosts: [excel]
tools: [create_chart, format_range]
mcpServers: [workiq]
---

Instructions here.`;

    const result = parseAgentFrontmatter(md);
    expect(result.metadata.name).toBe('Test Agent');
    expect(result.metadata.hosts).toEqual(['excel']);
    expect(result.metadata.defaultForHosts).toEqual(['excel']);
    expect(result.metadata.tools).toEqual(['create_chart', 'format_range']);
    expect(result.metadata.mcpServers).toEqual(['workiq']);
    expect(result.instructions).toContain('Instructions here.');
  });

  it('handles missing frontmatter and ignores invalid hosts', () => {
    expect(parseAgentFrontmatter('Instructions only.').metadata.name).toBe('unknown');

    const result = parseAgentFrontmatter(`---
name: Bad Host
description: desc
version: 1.0.0
hosts: [invalid_host]
defaultForHosts: []
---
Body`);
    expect(result.metadata.hosts).toEqual([]);
  });
});

describe('agentService — bundled agents', () => {
  it('returns bundled agents for each supported host', () => {
    for (const host of ['excel', 'powerpoint', 'word', 'outlook'] as const) {
      const agents = getAgents(host);
      expect(agents.length).toBeGreaterThan(0);
      expect(agents.every(a => a.metadata.hosts.includes(host))).toBe(true);
      expect(getDefaultAgent(host)?.metadata.hosts).toContain(host);
    }
  });

  it('returns no agents for unknown host', () => {
    expect(getAgents('unknown' as never)).toEqual([]);
    expect(getDefaultAgent('unknown' as never)).toBeUndefined();
  });

  it('does not surface imported/plugin agents in the picker', () => {
    expect(getPickerAgents('excel')).toEqual([]);
    expect(getPickerAgents('powerpoint')).toEqual([]);
    expect(getPickerAgents('word')).toEqual([]);
    expect(getPickerAgents('outlook')).toEqual([]);
  });

  it('looks up bundled agents and instructions by name', () => {
    const first = getAgents('excel')[0];
    expect(getAgent(first.metadata.name, 'excel')).toBeDefined();
    expect(getAgentInstructions(first.metadata.name, 'excel')).toBe(first.instructions);
    expect(getAllAgents().some(agent => agent.metadata.name === first.metadata.name)).toBe(true);
  });

  it('falls back to the host default for an invalid active agent', () => {
    expect(resolveActiveAgent('Missing Agent', 'excel')?.metadata.name).toBe('Excel');
  });
});
