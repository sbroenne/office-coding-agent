/**
 * Unit tests for the agent service: parseAgentFrontmatter, getAgents, getAgent, getAgentInstructions.
 *
 * These exercise the real `agentService` module which imports
 * bundled `.md` agent files via the rawMarkdownPlugin in vitest.config.ts.
 */

import { describe, it, expect } from 'vitest';
import {
  parseAgentFrontmatter,
  getAgents,
  getAgent,
  getAgentInstructions,
  getDefaultAgent,
  resolveActiveAgent,
} from '@/services/agents/agentService';

// ─── parseAgentFrontmatter ───

describe('parseAgentFrontmatter', () => {
  it('parses name, description, and version from YAML frontmatter', () => {
    const raw = `---
name: TestAgent
description: A test agent
version: 2.0.0
hosts: [excel]
defaultForHosts: [excel]
---

Some instructions here.`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.name).toBe('TestAgent');
    expect(result.metadata.description).toBe('A test agent');
    expect(result.metadata.version).toBe('2.0.0');
    expect(result.metadata.hosts).toEqual(['excel']);
    expect(result.metadata.defaultForHosts).toEqual(['excel']);
    expect(result.instructions).toBe('Some instructions here.');
  });

  it('parses host arrays in multiline YAML list format', () => {
    const raw = `---
name: HostTargeted
description: Agent for two hosts
version: 1.0.0
hosts:
  - excel
  - powerpoint
defaultForHosts:
  - powerpoint
---

Body content.`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.hosts).toEqual(['excel', 'powerpoint']);
    expect(result.metadata.defaultForHosts).toEqual(['powerpoint']);
  });

  it('drops invalid host values from host arrays', () => {
    const raw = `---
name: InvalidHosts
description: Agent with invalid host entries
version: 1.0.0
hosts: [excel, word, powerpoint]
defaultForHosts: [word, excel]
---

Body content.`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.hosts).toEqual(['excel', 'powerpoint']);
    expect(result.metadata.defaultForHosts).toEqual(['excel']);
  });

  it('handles multiline description with >', () => {
    const raw = `---
name: Multi
description: >
  This is a long
  multiline description
version: 1.0.0
---

Body content.`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.name).toBe('Multi');
    expect(result.metadata.description).toBe('This is a long multiline description');
    expect(result.instructions).toBe('Body content.');
  });

  it('returns defaults when no frontmatter present', () => {
    const raw = 'Just plain instructions.';
    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.name).toBe('unknown');
    expect(result.metadata.description).toBe('');
    expect(result.metadata.version).toBe('0.0.0');
    expect(result.metadata.hosts).toEqual([]);
    expect(result.metadata.defaultForHosts).toEqual([]);
    expect(result.instructions).toBe('Just plain instructions.');
  });

  it('returns defaults when frontmatter is not closed', () => {
    const raw = `---
name: Unclosed
description: Missing end marker`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.name).toBe('unknown');
  });

  it('handles empty body after frontmatter', () => {
    const raw = `---
name: EmptyBody
description: No instructions
version: 1.0.0
---`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.name).toBe('EmptyBody');
    expect(result.instructions).toBe('');
  });

  it('handles leading whitespace before frontmatter', () => {
    const raw = `
---
name: Padded
description: Has leading whitespace
version: 0.1.0
---

Instructions.`;

    const result = parseAgentFrontmatter(raw);
    expect(result.metadata.name).toBe('Padded');
    expect(result.instructions).toBe('Instructions.');
  });
});

// ─── Bundled agents ───

describe('getAgents', () => {
  it('returns at least one bundled agent', () => {
    const agents = getAgents('excel');
    expect(agents.length).toBeGreaterThanOrEqual(1);
  });

  it('the default Excel agent is present', () => {
    const agents = getAgents('excel');
    const excel = agents.find(a => a.metadata.name === 'Excel');
    expect(excel).toBeDefined();
    expect(excel!.metadata.description).toBeTruthy();
    expect(excel!.instructions).toContain('Excel');
    expect(excel!.metadata.hosts).toContain('excel');
  });

  it('returns no agents for PowerPoint until one is added', () => {
    expect(getAgents('powerpoint')).toEqual([]);
  });
});

describe('getAgent', () => {
  it('returns the Excel agent by name', () => {
    const agent = getAgent('Excel');
    expect(agent).toBeDefined();
    expect(agent!.metadata.name).toBe('Excel');
  });

  it('returns undefined for nonexistent agent', () => {
    expect(getAgent('NonExistent')).toBeUndefined();
  });
});

describe('getAgentInstructions', () => {
  it('returns instructions for the Excel agent', () => {
    const instructions = getAgentInstructions('Excel');
    expect(instructions.length).toBeGreaterThan(0);
    expect(instructions).toContain('Discover first');
  });

  it('includes execution rules adapted from MCP skill', () => {
    const instructions = getAgentInstructions('Excel');
    expect(instructions).toContain('Never Ask Clarifying Questions');
    expect(instructions).toContain('Always End With a Text Summary');
    expect(instructions).toContain('Format Data Professionally');
    expect(instructions).toContain('Tool Selection Quick Reference');
  });

  it('returns empty string for unknown agent', () => {
    expect(getAgentInstructions('DoesNotExist')).toBe('');
  });
});

describe('default and active resolution', () => {
  it('returns Excel as default agent for excel host', () => {
    const defaultAgent = getDefaultAgent('excel');
    expect(defaultAgent).toBeDefined();
    expect(defaultAgent!.metadata.name).toBe('Excel');
  });

  it('returns undefined default for host with no configured agents', () => {
    expect(getDefaultAgent('powerpoint')).toBeUndefined();
  });

  it('falls back to host default when active agent does not match host', () => {
    const resolved = resolveActiveAgent('NonExistent', 'excel');
    expect(resolved).toBeDefined();
    expect(resolved!.metadata.name).toBe('Excel');
  });
});
