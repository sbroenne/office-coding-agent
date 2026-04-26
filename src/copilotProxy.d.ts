declare module '*/copilotProxy.mjs' {
  export function getRegisteredToolNames(
    toolDefs?: Array<{ name?: string }>,
    availableTools?: string[]
  ): string[];

  export function applySessionToolAccessToCustomAgents(
    customAgents?: Array<{ tools?: string[] | null } & Record<string, unknown>>,
    sessionToolNames?: string[]
  ): Array<{ tools?: string[] } & Record<string, unknown>>;

  export function mergePluginMcpServers(
    baseServers?: Record<string, unknown>,
    pluginServers?: Record<string, unknown>
  ): Record<string, unknown>;
}
