import { describe, it, expect } from 'vitest';
import {
  parseAgentFrontmatter,
  getAgents,
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
---

Instructions here.`;

    const result = parseAgentFrontmatter(md);
    expect(result.metadata.name).toBe('Test Agent');
    expect(result.metadata.hosts).toEqual(['excel']);
    expect(result.metadata.defaultForHosts).toEqual(['excel']);
    expect(result.instructions).toContain('Instructions here.');
  });

  it('silently handles missing required fields', () => {
    const md = `---
name: Incomplete Agent
---
No version or hosts.`;

    const result = parseAgentFrontmatter(md);
    expect(result.metadata.name).toBe('Incomplete Agent');
    expect(result.metadata.hosts).toEqual([]);
  });

  it('silently ignores invalid host values', () => {
    const md = `---
name: Bad Host
description: desc
version: 1.0.0
hosts: [invalid_host]
defaultForHosts: []
---
Body`;

    const result = parseAgentFrontmatter(md);
    expect(result.metadata.hosts).toEqual([]);
  });
});

describe('agentService — getAgents', () => {
  it('returns an array of agents', () => {
    const agents = getAgents();
    expect(Array.isArray(agents)).toBe(true);
    expect(agents.length).toBeGreaterThan(0);
  });

  it('every agent has a name and at least one host', () => {
    for (const agent of getAgents()) {
      expect(agent.metadata.name.length).toBeGreaterThan(0);
      expect(agent.metadata.hosts.length).toBeGreaterThan(0);
    }
  });

  it('filters by host', () => {
    const excelAgents = getAgents('excel');
    expect(excelAgents.every(a => a.metadata.hosts.includes('excel'))).toBe(true);
  });

  it('returns empty array for unknown host', () => {
    expect(getAgents('unknown' as never)).toEqual([]);
  });
});

describe('agentService — getAgent', () => {
  it('returns an agent by name', () => {
    const agents = getAgents();
    const first = agents[0];
    const found = getAgent(first.metadata.name);
    expect(found).toBeDefined();
    expect(found?.metadata.name).toBe(first.metadata.name);
  });

  it('returns undefined for unknown name', () => {
    expect(getAgent('NonExistent__xyz')).toBeUndefined();
  });
});

describe('agentService — getAgentInstructions', () => {
  it('returns instructions string for a known agent', () => {
    const agents = getAgents();
    const name = agents[0].metadata.name;
    const instructions = getAgentInstructions(name);
    expect(typeof instructions).toBe('string');
    expect(instructions.length).toBeGreaterThan(0);
  });

  it('returns empty string for unknown agent', () => {
    expect(getAgentInstructions('NoSuchAgent')).toBe('');
  });
});

describe('agentService — getDefaultAgent', () => {
  it('returns the default Excel agent', () => {
    const agent = getDefaultAgent('excel');
    expect(agent).toBeDefined();
    expect(agent?.metadata.hosts).toContain('excel');
  });

  it('returns undefined for unknown host', () => {
    expect(getDefaultAgent('unknown' as never)).toBeUndefined();
  });
});

describe('agentService — resolveActiveAgent', () => {
  it('returns the named agent when valid for the host', () => {
    const agents = getAgents('excel');
    const name = agents[0].metadata.name;
    const resolved = resolveActiveAgent(name, 'excel');
    expect(resolved).toBeDefined();
    expect(resolved?.metadata.name).toBe(name);
  });

  it('falls back to default when activeAgentId is empty string', () => {
    const resolved = resolveActiveAgent('', 'excel');
    expect(resolved).toBeDefined();
  });

  it('falls back to default when agent not valid for host', () => {
    const resolved = resolveActiveAgent('NonExistent__xyz', 'excel');
    expect(resolved).toBeDefined();
  });
});
