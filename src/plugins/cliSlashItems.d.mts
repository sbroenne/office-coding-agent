export interface CliSlashItem {
  type: 'skill' | 'prompt';
  name: string;
  description: string;
  plugin?: string;
}

export interface CliSlashItems {
  skills: CliSlashItem[];
}

export interface CliSlashItemsOptions {
  installedPluginsDir?: string;
}

export function getCliSlashItems(options?: CliSlashItemsOptions): Promise<CliSlashItems>;
