export interface RegisteredMarketplaceEntry {
  slug: string;
  name: string;
  source: string;
  isBuiltIn: boolean;
  isOwn: boolean;
  registeredKey: string;
  pluginCount: number;
}

export interface RemoveResult {
  success: boolean;
  message: string;
}

export declare const OCA_MARKETPLACE_KEY: string;
export declare const BUILTIN_KEYS: string[];

export declare function readConfig(configPath: string): Record<string, unknown>;
export declare function findMarketplaceManifest(cacheDir: string): Record<string, unknown> | null;
export declare function repoCacheSlugs(repo: string): string[];
export declare function listMarketplaces(cacheDir: string, configPath: string): RegisteredMarketplaceEntry[];
export declare function removeMarketplace(registeredKey: string): RemoveResult;
