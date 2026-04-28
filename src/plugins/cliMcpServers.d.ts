declare module '*/cliMcpServers.mjs' {
  import type { McpServerConfig } from '@/types';

  export interface CopilotCommandResult {
    success: boolean;
    stdout: string;
    stderr: string;
    message?: string;
    code?: number | null;
  }

  export interface CopilotMcpServerListResult {
    servers: McpServerConfig[];
    error?: string;
  }

  export interface CopilotMcpServerOptions {
    command?: string;
    timeoutMs?: number;
    runCommand?: (options: CopilotMcpServerOptions) => Promise<CopilotCommandResult>;
  }

  export function runCopilotMcpListCommand(
    options?: CopilotMcpServerOptions
  ): Promise<CopilotCommandResult>;
  export function parseCopilotMcpListJson(stdout: string): McpServerConfig[];
  export function getCliMcpServers(
    options?: CopilotMcpServerOptions
  ): Promise<CopilotMcpServerListResult>;
}
