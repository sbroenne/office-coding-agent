/**
 * Management tools for installing, removing, and toggling skills, agents, and MCP servers.
 *
 * These are general-purpose tools — not tied to any Office host.
 * They call into the Zustand settings store to persist changes.
 * Changes take effect on the next conversation (session restart).
 */

import type { Tool, ToolInvocation, ToolResultObject } from '@github/copilot-sdk';
import { useSettingsStore } from '@/stores';
import { getSkills } from '@/services/skills';
import { getAllAgents } from '@/services/agents';
import type { AgentHost } from '@/types/agent';

// ─── manage_skills ────────────────────────────────────────────────────────────

export const manageSkillsTool: Tool = {
  name: 'manage_skills',
  description:
    'Manage agent skills. Actions: "list" (show all skills and active state), "install" (add a new skill from structured data), "install_from_npm" (register a skillpm skill package from npm — see skillpm.dev — whose SKILL.md files are loaded at session start), "remove" (delete an imported skill by name), "remove_npm_package" (unregister a skillpm package by name), "toggle" (enable/disable a skill by name). Packages must follow the skillpm spec: skills/<name>/SKILL.md inside the package.',
  parameters: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['list', 'install', 'install_from_npm', 'remove', 'remove_npm_package', 'toggle'],
      },
      // install params
      name: { type: 'string', description: 'Skill name (install, remove, toggle).' },
      description: { type: 'string', description: 'Skill description (install).' },
      version: { type: 'string', description: 'Semantic version, e.g. "1.0.0" (install).' },
      hosts: {
        type: 'array',
        items: { type: 'string' },
        description:
          'Office hosts where this skill is available: "excel", "powerpoint", "word". Empty array = all hosts (install).',
      },
      tags: {
        type: 'array',
        items: { type: 'string' },
        description: 'Tags for categorization (install).',
      },
      content: {
        type: 'string',
        description: 'Markdown body of the skill — injected into the system prompt (install).',
      },
      // install_from_npm / remove_npm_package params
      package: {
        type: 'string',
        description:
          'skillpm package name from npmjs.org (install_from_npm, remove_npm_package), e.g. "skillpm-skill" or "@myorg/my-skills". The package must follow the skillpm spec (skills/<name>/SKILL.md). See skillpm.dev for how to create and publish skill packages.',
      },
    },
    required: ['action'],
  },
  handler: (args: unknown, _invocation: ToolInvocation): ToolResultObject | string => {
    const { action, name, description, version, hosts, tags, content } = args as {
      action: string;
      name?: string;
      description?: string;
      version?: string;
      hosts?: string[];
      tags?: string[];
      content?: string;
      package?: string;
    };

    const store = useSettingsStore.getState();

    if (action === 'list') {
      const skills = getSkills();
      const activeNames = store.activeSkillNames;
      return JSON.stringify({
        skills: skills.map(s => ({
          name: s.metadata.name,
          description: s.metadata.description,
          version: s.metadata.version,
          hosts: s.metadata.hosts as string[],
          tags: s.metadata.tags,
          active: activeNames === null ? true : activeNames.includes(s.metadata.name),
        })),
        npmSkillPackages: store.npmSkillPackages,
        count: skills.length,
      });
    }

    if (action === 'install') {
      if (!name) return JSON.stringify({ error: 'name is required for install' });
      if (!content) return JSON.stringify({ error: 'content is required for install' });

      const skill = {
        metadata: {
          name,
          description: description ?? '',
          version: version ?? '1.0.0',
          hosts: (hosts ?? []) as AgentHost[],
          tags: tags ?? [],
        },
        content,
      };
      store.importSkills([skill]);
      return JSON.stringify({
        installed: true,
        name: skill.metadata.name,
        message: 'Skill installed. It will be active in the next conversation.',
      });
    }

    if (action === 'install_from_npm') {
      const { package: packageName } = args as { package?: string };
      if (!packageName)
        return JSON.stringify({ error: 'package is required for install_from_npm' });

      store.addNpmSkillPackage(packageName);
      return JSON.stringify({
        registered: true,
        package: packageName,
        message:
          `skillpm package "${packageName}" registered. ` +
          'The proxy will install it via npm and load all skills/<name>/SKILL.md files at the start of the next conversation. ' +
          'See skillpm.dev for how to create and publish skill packages.',
      });
    }

    if (action === 'remove_npm_package') {
      const { package: packageName } = args as { package?: string };
      if (!packageName)
        return JSON.stringify({ error: 'package is required for remove_npm_package' });

      store.removeNpmSkillPackage(packageName);
      return JSON.stringify({ removed: true, package: packageName });
    }

    if (action === 'remove') {
      if (!name) return JSON.stringify({ error: 'name is required for remove' });
      store.removeImportedSkill(name);
      return JSON.stringify({ removed: true, name });
    }

    if (action === 'toggle') {
      if (!name) return JSON.stringify({ error: 'name is required for toggle' });
      store.toggleSkill(name);
      const activeNames = useSettingsStore.getState().activeSkillNames;
      const isActive = activeNames === null ? true : activeNames.includes(name);
      return JSON.stringify({ toggled: true, name, active: isActive });
    }

    return JSON.stringify({ error: `Unknown action: ${action}` });
  },
};

// ─── manage_agents ────────────────────────────────────────────────────────────

export const manageAgentsTool: Tool = {
  name: 'manage_agents',
  description:
    'Manage custom agents. Actions: "list" (show all agents and active state), "install" (add a new agent from structured data), "remove" (delete an imported agent by name), "set_active" (switch the active agent by name).',
  parameters: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['list', 'install', 'remove', 'set_active'],
      },
      // install / remove / set_active params
      name: { type: 'string', description: 'Agent name (install, remove, set_active).' },
      description: { type: 'string', description: 'Agent description (install).' },
      version: { type: 'string', description: 'Semantic version, e.g. "1.0.0" (install).' },
      hosts: {
        type: 'array',
        items: { type: 'string' },
        description:
          'Office hosts where this agent is available: "excel", "powerpoint", "word" (install). Required.',
      },
      defaultForHosts: {
        type: 'array',
        items: { type: 'string' },
        description: 'Hosts where this agent is the default choice (install).',
      },
      instructions: {
        type: 'string',
        description:
          'Markdown body of the agent — injected into the system prompt as agent-specific instructions (install).',
      },
      tools: {
        type: 'array',
        items: { type: 'string' },
        description:
          'Allowlist of built-in tool names available in this agent. Omit = all tools (install).',
      },
      mcpServers: {
        type: 'array',
        items: { type: 'string' },
        description:
          'Allowlist of MCP server names available in this agent. Omit = all servers (install).',
      },
    },
    required: ['action'],
  },
  handler: (args: unknown, _invocation: ToolInvocation): ToolResultObject | string => {
    const {
      action,
      name,
      description,
      version,
      hosts,
      defaultForHosts,
      instructions,
      tools,
      mcpServers,
    } = args as {
      action: string;
      name?: string;
      description?: string;
      version?: string;
      hosts?: string[];
      defaultForHosts?: string[];
      instructions?: string;
      tools?: string[];
      mcpServers?: string[];
    };

    const store = useSettingsStore.getState();

    if (action === 'list') {
      const agents = getAllAgents();
      const activeAgentId = store.activeAgentId;
      return JSON.stringify({
        agents: agents.map(a => ({
          name: a.metadata.name,
          description: a.metadata.description,
          version: a.metadata.version,
          hosts: a.metadata.hosts,
          defaultForHosts: a.metadata.defaultForHosts,
          tools: a.metadata.tools,
          mcpServers: a.metadata.mcpServers,
          active: a.metadata.name === activeAgentId,
        })),
        activeAgentId,
        count: agents.length,
      });
    }

    if (action === 'install') {
      if (!name) return JSON.stringify({ error: 'name is required for install' });
      if (!hosts || hosts.length === 0)
        return JSON.stringify({ error: 'hosts is required for install (at least one host)' });
      if (!instructions) return JSON.stringify({ error: 'instructions is required for install' });

      const agent = {
        metadata: {
          name,
          description: description ?? '',
          version: version ?? '1.0.0',
          hosts: hosts as AgentHost[],
          defaultForHosts: (defaultForHosts ?? []) as AgentHost[],
          tools,
          mcpServers,
        },
        instructions,
      };
      store.importAgents([agent]);
      return JSON.stringify({
        installed: true,
        name: agent.metadata.name,
        message:
          'Agent installed. Use set_active to switch to it, or it will be available in the agent picker.',
      });
    }

    if (action === 'remove') {
      if (!name) return JSON.stringify({ error: 'name is required for remove' });
      store.removeImportedAgent(name);
      return JSON.stringify({ removed: true, name });
    }

    if (action === 'set_active') {
      if (!name) return JSON.stringify({ error: 'name is required for set_active' });
      store.setActiveAgent(name);
      const currentActive = useSettingsStore.getState().activeAgentId;
      return JSON.stringify({
        activeAgentId: currentActive,
        message:
          currentActive === name
            ? `Active agent set to "${name}". Changes take effect in the next conversation.`
            : `Agent "${name}" not found or not available for the current host. Active agent remains "${currentActive}".`,
      });
    }

    return JSON.stringify({ error: `Unknown action: ${action}` });
  },
};

// ─── manage_mcp_servers ───────────────────────────────────────────────────────

export const manageMcpServersTool: Tool = {
  name: 'manage_mcp_servers',
  description:
    'Manage MCP (Model Context Protocol) servers. Actions: "list" (show all servers and active state), "install" (add a new MCP server endpoint), "remove" (delete an MCP server by name), "toggle" (enable/disable an MCP server by name).',
  parameters: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        description: 'Operation to perform',
        enum: ['list', 'install', 'remove', 'toggle'],
      },
      name: { type: 'string', description: 'Server name (install, remove, toggle).' },
      description: { type: 'string', description: 'Server description (install).' },
      transport: {
        type: 'string',
        description:
          'Transport protocol (install). Use "http" or "sse" for remote servers, "stdio" for local subprocess servers (e.g. npx). Default: "http".',
        enum: ['http', 'sse', 'stdio'],
      },
      // http/sse fields
      url: {
        type: 'string',
        description: 'MCP server endpoint URL — required for http/sse transport (install).',
      },
      headers: {
        type: 'object',
        additionalProperties: { type: 'string' },
        description:
          'Optional HTTP headers, e.g. Authorization — for http/sse transport (install).',
      },
      // stdio fields
      command: {
        type: 'string',
        description: 'Executable command — required for stdio transport, e.g. "npx" (install).',
      },
      args: {
        type: 'array',
        items: { type: 'string' },
        description:
          'Command arguments — for stdio transport, e.g. ["-y", "@some/mcp-server"] (install).',
      },
      env: {
        type: 'object',
        additionalProperties: { type: 'string' },
        description:
          'Optional environment variables to pass to the subprocess — for stdio transport (install).',
      },
    },
    required: ['action'],
  },
  handler: (args: unknown, _invocation: ToolInvocation): ToolResultObject | string => {
    const {
      action,
      name,
      description,
      url,
      transport,
      headers,
      command,
      args: cmdArgs,
      env,
    } = args as {
      action: string;
      name?: string;
      description?: string;
      url?: string;
      transport?: 'http' | 'sse' | 'stdio';
      headers?: Record<string, string>;
      command?: string;
      args?: string[];
      env?: Record<string, string>;
    };

    const store = useSettingsStore.getState();

    if (action === 'list') {
      const servers = store.importedMcpServers;
      const activeNames = store.activeMcpServerNames;
      return JSON.stringify({
        servers: servers.map(s => ({
          name: s.name,
          description: s.description,
          transport: s.transport,
          ...(s.transport === 'stdio' ? { command: s.command, args: s.args } : { url: s.url }),
          active: activeNames === null ? true : activeNames.includes(s.name),
        })),
        count: servers.length,
      });
    }

    if (action === 'install') {
      if (!name) return JSON.stringify({ error: 'name is required for install' });

      const resolvedTransport = transport ?? 'http';

      if (resolvedTransport === 'stdio') {
        if (!command) return JSON.stringify({ error: 'command is required for stdio transport' });

        const server = {
          name,
          description,
          transport: 'stdio' as const,
          command,
          args: cmdArgs ?? [],
          env,
        };
        store.importMcpServers([server]);
        return JSON.stringify({
          installed: true,
          name: server.name,
          message: 'MCP server installed. It will be available in the next conversation.',
        });
      }

      // http / sse
      if (!url) return JSON.stringify({ error: 'url is required for http/sse transport' });

      const server = {
        name,
        description,
        url,
        transport: resolvedTransport,
        headers,
      };
      store.importMcpServers([server]);
      return JSON.stringify({
        installed: true,
        name: server.name,
        message: 'MCP server installed. It will be available in the next conversation.',
      });
    }

    if (action === 'remove') {
      if (!name) return JSON.stringify({ error: 'name is required for remove' });
      store.removeMcpServer(name);
      return JSON.stringify({ removed: true, name });
    }

    if (action === 'toggle') {
      if (!name) return JSON.stringify({ error: 'name is required for toggle' });
      store.toggleMcpServer(name);
      const activeNames = useSettingsStore.getState().activeMcpServerNames;
      const isActive = activeNames === null ? true : activeNames.includes(name);
      return JSON.stringify({ toggled: true, name, active: isActive });
    }

    return JSON.stringify({ error: `Unknown action: ${action}` });
  },
};

/** All management tools — included for every host */
export const managementTools: Tool[] = [manageSkillsTool, manageAgentsTool, manageMcpServersTool];
