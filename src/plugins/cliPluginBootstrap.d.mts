export interface CopilotPluginCommandResult {
  success: boolean;
  stdout: string;
  stderr?: string;
  message?: string;
  code?: number | null;
}

export interface OfficeCliPluginMarketplace {
  name: string;
  source: string;
}

export const OFFICE_CODING_AGENT_MARKETPLACE: OfficeCliPluginMarketplace;
export const REQUIRED_OFFICE_PLUGINS: string[];
export const REQUIRED_OFFICE_PLUGIN_SPECS: string[];

export function copilotHomeDir(): string;

export function getInstalledOfficePluginDirectories(options?: {
  home?: string;
  marketplace?: OfficeCliPluginMarketplace;
  plugins?: string[];
  fileExists?: (dir: string) => boolean;
}): string[];

export function runCopilotPluginCommand(
  args: string[],
  options?: { timeoutMs?: number; command?: string }
): Promise<CopilotPluginCommandResult>;

export function hasMarketplace(
  listOutput: string,
  marketplace?: OfficeCliPluginMarketplace
): boolean;

export function installedPluginSpecs(listOutput: string): Set<string>;

export function ensureOfficeCliPlugins(options?: {
  runCommand?: (args: string[]) => Promise<CopilotPluginCommandResult>;
  logger?: Pick<Console, 'log' | 'warn'>;
  marketplace?: OfficeCliPluginMarketplace;
  pluginSpecs?: string[];
}): Promise<{ success: boolean; error?: unknown }>;
