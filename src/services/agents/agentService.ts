import type { AgentConfig, AgentHost, AgentMetadata } from '@/types/agent';
import type { OfficeHostApp } from '@/services/office/host';

// Import bundled agent files as raw strings (via Vite md-raw plugin)
import excelAgentRaw from '@/agents/excel/AGENT.md';

/**
 * Parse YAML frontmatter from an agent markdown file.
 * Handles the `---` delimited block at the top of the file.
 */
export function parseAgentFrontmatter(raw: string): AgentConfig {
  const trimmed = raw.trimStart();

  if (!trimmed.startsWith('---')) {
    return {
      metadata: {
        name: 'unknown',
        description: '',
        version: '0.0.0',
        hosts: [],
        defaultForHosts: [],
      },
      instructions: trimmed,
    };
  }

  const endIndex = trimmed.indexOf('---', 3);
  if (endIndex === -1) {
    return {
      metadata: {
        name: 'unknown',
        description: '',
        version: '0.0.0',
        hosts: [],
        defaultForHosts: [],
      },
      instructions: trimmed,
    };
  }

  const yamlBlock = trimmed.slice(3, endIndex).trim();
  const instructions = trimmed.slice(endIndex + 3).trim();

  // Simple YAML parser for flat key-value pairs
  const metadata: AgentMetadata = {
    name: '',
    description: '',
    version: '0.0.0',
    hosts: [],
    defaultForHosts: [],
  };

  let currentKey = '';
  let isMultilineValue = false;
  let multilineValue = '';

  for (const line of yamlBlock.split('\n')) {
    const trimmedLine = line.trim();

    if (
      trimmedLine.startsWith('- ') &&
      (currentKey === 'hosts' || currentKey === 'defaultForHosts')
    ) {
      setAgentArrayField(metadata, currentKey, [trimmedLine.slice(2).trim()]);
      continue;
    }

    // Multiline continuation (indented lines after "key: >")
    if (isMultilineValue && (line.startsWith('  ') || line.startsWith('\t'))) {
      multilineValue += (multilineValue ? ' ' : '') + trimmedLine;
      continue;
    }

    // Flush multiline value
    if (isMultilineValue) {
      setAgentField(metadata, currentKey, multilineValue);
      isMultilineValue = false;
      multilineValue = '';
    }

    // Key-value pairs
    const colonIndex = trimmedLine.indexOf(':');
    if (colonIndex === -1) continue;

    currentKey = trimmedLine.slice(0, colonIndex).trim();
    const value = trimmedLine.slice(colonIndex + 1).trim();

    if (value === '>' || value === '|') {
      isMultilineValue = true;
      multilineValue = '';
    } else if (value === '') {
      continue;
    } else {
      if (currentKey === 'hosts' || currentKey === 'defaultForHosts') {
        setAgentArrayField(metadata, currentKey, parseInlineArray(value));
        continue;
      }
      setAgentField(metadata, currentKey, value);
    }
  }

  // Flush any trailing multiline value
  if (isMultilineValue && multilineValue) {
    setAgentField(metadata, currentKey, multilineValue);
  }

  return { metadata, instructions };
}

function setAgentField(metadata: AgentMetadata, key: string, value: string): void {
  switch (key) {
    case 'name':
      metadata.name = value;
      break;
    case 'description':
      metadata.description = value;
      break;
    case 'version':
      metadata.version = value;
      break;
  }
}

function parseInlineArray(value: string): string[] {
  const trimmed = value.trim();
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    return trimmed
      .slice(1, -1)
      .split(',')
      .map(item => item.trim())
      .filter(Boolean);
  }
  return [trimmed];
}

const SUPPORTED_AGENT_HOSTS: AgentHost[] = ['excel', 'powerpoint'];

function isAgentHost(value: string): value is AgentHost {
  return SUPPORTED_AGENT_HOSTS.includes(value as AgentHost);
}

function setAgentArrayField(metadata: AgentMetadata, key: string, values: string[]): void {
  const normalized = values.map(v => v.toLowerCase()).filter(isAgentHost);

  if (key === 'hosts') {
    metadata.hosts = Array.from(new Set([...metadata.hosts, ...normalized]));
  }

  if (key === 'defaultForHosts') {
    metadata.defaultForHosts = Array.from(new Set([...metadata.defaultForHosts, ...normalized]));
  }
}

/** All bundled agents, parsed at module load time. */
const bundledAgents: AgentConfig[] = [parseAgentFrontmatter(excelAgentRaw)];
let importedAgents: AgentConfig[] = [];

export function getBundledAgents(): AgentConfig[] {
  return bundledAgents;
}

export function getImportedAgents(): AgentConfig[] {
  return importedAgents;
}

export function setImportedAgents(agents: AgentConfig[]): void {
  importedAgents = agents;
}

function toAgentHost(host: OfficeHostApp): AgentHost | undefined {
  if (host === 'excel' || host === 'powerpoint') return host;
  return undefined;
}

/**
 * Get all loaded agents.
 */
export function getAgents(host: OfficeHostApp = 'excel'): AgentConfig[] {
  const targetHost = toAgentHost(host);
  if (!targetHost) return [];
  return [...bundledAgents, ...importedAgents].filter(agent =>
    agent.metadata.hosts.includes(targetHost)
  );
}

export function getAllAgents(): AgentConfig[] {
  return [...bundledAgents, ...importedAgents];
}

/**
 * Get a single agent by name.
 */
export function getAgent(name: string, host: OfficeHostApp = 'excel'): AgentConfig | undefined {
  return getAgents(host).find(a => a.metadata.name === name);
}

/**
 * Get the instructions for a specific agent by name.
 * Returns an empty string if the agent is not found.
 */
export function getAgentInstructions(name: string, host: OfficeHostApp = 'excel'): string {
  return getAgent(name, host)?.instructions ?? '';
}

export function getDefaultAgent(host: OfficeHostApp = 'excel'): AgentConfig | undefined {
  const targetHost = toAgentHost(host);
  if (!targetHost) return undefined;

  const hostAgents = getAgents(targetHost);
  return (
    hostAgents.find(agent => agent.metadata.defaultForHosts.includes(targetHost)) ?? hostAgents[0]
  );
}

export function resolveActiveAgent(
  activeAgentId: string,
  host: OfficeHostApp = 'excel'
): AgentConfig | undefined {
  return getAgent(activeAgentId, host) ?? getDefaultAgent(host);
}
