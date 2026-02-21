export type AgentHost = 'excel' | 'powerpoint' | 'word';

/** Parsed metadata from an agent's YAML frontmatter. */
export interface AgentMetadata {
  /** Display name shown in the agent picker */
  name: string;
  /** Brief description shown as secondary text */
  description: string;
  /** Semantic version (e.g. "1.0.0") */
  version: string;
  /** Office hosts where this agent is available. */
  hosts: AgentHost[];
  /** Hosts where this agent should be used as the default choice. */
  defaultForHosts: AgentHost[];
}

/** A loaded agent configuration with metadata and instructions. */
export interface AgentConfig {
  /** Parsed YAML frontmatter */
  metadata: AgentMetadata;
  /** Markdown body — injected into system prompt as agent-specific instructions */
  instructions: string;
}
