import { describe, it, expect, afterEach } from 'vitest';
import {
  parseAgentFrontmatter,
  getAgents,
  getAllAgents,
  getAgent,
  getAgentInstructions,
  getDefaultAgent,
  resolveActiveAgent,
  setImportedAgents,
} from '@/services/agents/agentService';
import type { AgentConfig } from '@/types/agent';

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

describe('agentService — parseAgentFrontmatter tools/mcpServers', () => {
  it('parses inline tools array', () => {
    const md = `---
name: Scoped Agent
description: desc
version: 1.0.0
hosts: [excel]
defaultForHosts: []
tools: [create_chart, format_range]
---
Instructions`;
    const result = parseAgentFrontmatter(md);
    expect(result.metadata.tools).toEqual(['create_chart', 'format_range']);
  });

  it('parses inline mcpServers array', () => {
    const md = `---
name: MCP Agent
description: desc
version: 1.0.0
hosts: [excel]
defaultForHosts: []
mcpServers: [my-server, other-server]
---
Instructions`;
    const result = parseAgentFrontmatter(md);
    expect(result.metadata.mcpServers).toEqual(['my-server', 'other-server']);
  });

  it('parses block-style tools list', () => {
    const md = `---
name: Block Agent
description: desc
version: 1.0.0
hosts: [excel]
defaultForHosts: []
tools:
  - create_chart
  - delete_sheet
---
Instructions`;
    const result = parseAgentFrontmatter(md);
    expect(result.metadata.tools).toEqual(['create_chart', 'delete_sheet']);
  });

  it('leaves tools undefined when not specified', () => {
    const md = `---
name: Plain Agent
description: desc
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---
Instructions`;
    const result = parseAgentFrontmatter(md);
    expect(result.metadata.tools).toBeUndefined();
    expect(result.metadata.mcpServers).toBeUndefined();
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

// ─── Imported agent visibility ────────────────────────────────────────────────

describe('agentService — imported agents', () => {
  const importedAgent: AgentConfig = {
    metadata: {
      name: 'Imported Excel Agent',
      description: 'Imported custom agent',
      version: '1.0.0',
      hosts: ['excel'],
      defaultForHosts: [],
    },
    instructions: 'Imported instructions.',
  };

  const pptOnlyAgent: AgentConfig = {
    metadata: {
      name: 'PPT Only Agent',
      description: 'Only for PowerPoint',
      version: '1.0.0',
      hosts: ['powerpoint'],
      defaultForHosts: [],
    },
    instructions: 'PPT instructions.',
  };

  afterEach(() => {
    setImportedAgents([]);
  });

  it('getAllAgents includes both bundled and imported agents', () => {
    setImportedAgents([importedAgent]);

    const all = getAllAgents();
    expect(all.some(a => a.metadata.name === 'Imported Excel Agent')).toBe(true);
    expect(all.some(a => a.metadata.name === 'Excel')).toBe(true);
  });

  it('getAllAgents returns only bundled agents when no imported agents are set', () => {
    const all = getAllAgents();
    expect(all.some(a => a.metadata.name === 'Excel')).toBe(true);
    expect(all.some(a => a.metadata.name === 'Imported Excel Agent')).toBe(false);
  });

  it('getAgents includes imported agent for matching host', () => {
    setImportedAgents([importedAgent]);

    const excelAgents = getAgents('excel');
    expect(excelAgents.some(a => a.metadata.name === 'Imported Excel Agent')).toBe(true);
  });

  it('getAgents excludes imported agent when host does not match', () => {
    setImportedAgents([pptOnlyAgent]);

    const excelAgents = getAgents('excel');
    expect(excelAgents.some(a => a.metadata.name === 'PPT Only Agent')).toBe(false);
  });

  it('getAgent finds an imported agent by name', () => {
    setImportedAgents([importedAgent]);

    const found = getAgent('Imported Excel Agent', 'excel');
    expect(found).toBeDefined();
    expect(found?.metadata.name).toBe('Imported Excel Agent');
    expect(found?.instructions).toBe('Imported instructions.');
  });

  it('getAgentInstructions returns instructions for an imported agent', () => {
    setImportedAgents([importedAgent]);

    const instructions = getAgentInstructions('Imported Excel Agent', 'excel');
    expect(instructions).toBe('Imported instructions.');
  });

  it('resolveActiveAgent returns imported agent when it matches the activeAgentId', () => {
    setImportedAgents([importedAgent]);

    const resolved = resolveActiveAgent('Imported Excel Agent', 'excel');
    expect(resolved?.metadata.name).toBe('Imported Excel Agent');
  });

  it('resolveActiveAgent falls back to default when imported agent is not for current host', () => {
    setImportedAgents([pptOnlyAgent]);

    // PPT Only Agent is not valid for excel
    const resolved = resolveActiveAgent('PPT Only Agent', 'excel');
    expect(resolved?.metadata.name).toBe('Excel'); // falls back to default Excel agent
  });
});
