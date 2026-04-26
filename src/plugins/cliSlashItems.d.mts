export interface CliSlashItem {
  type: 'skill' | 'prompt';
  name: string;
  description: string;
  plugin?: string;
  source?: string;
}

export interface CliSlashItems {
  skills: CliSlashItem[];
  prompts: CliSlashItem[];
}

export interface CliSlashItemsOptions {
  installedPluginsDir?: string;
  workspacePromptsDir?: string;
}

export function getCliSlashItems(options?: CliSlashItemsOptions): Promise<CliSlashItems>;
