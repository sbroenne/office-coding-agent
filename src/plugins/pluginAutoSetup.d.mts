export const OCA_MARKETPLACE: string;
export const OCA_MARKETPLACE_NAME: string;
export const HOST_PLUGINS: string[];

export interface PluginAutoSetupDependencies {
  getPluginManagerFn?: () => Promise<{
    refresh: () => Promise<void>;
    list: () => Array<{ name: string }>;
  }>;
  listPluginMarketplacesFn?: () => Promise<{
    success: boolean;
    data?: Array<{ name: string }>;
  }>;
  addPluginMarketplaceFn?: (source: string) => Promise<{ success: boolean; error?: string }>;
  installPluginFn?: (spec: string) => Promise<{ success: boolean; error?: string }>;
  logger?: Pick<Console, 'log' | 'warn'>;
}

export function autoSetupPlugins(deps?: PluginAutoSetupDependencies): Promise<void>;
