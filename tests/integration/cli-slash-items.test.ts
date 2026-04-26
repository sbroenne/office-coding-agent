import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { getCliSlashItems } from '@/../src/plugins/cliSlashItems.mjs';

const tempDirs: string[] = [];

async function makeInstalledPluginDir() {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'office-cli-slash-items-'));
  tempDirs.push(dir);
  return dir;
}

describe('CLI slash items', () => {
  afterEach(async () => {
    await Promise.all(tempDirs.splice(0).map(dir => fs.rm(dir, { recursive: true, force: true })));
  });

  it('discovers skills and prompts from installed Copilot CLI plugins', async () => {
    const installedPluginsDir = await makeInstalledPluginDir();
    const pluginRoot = path.join(installedPluginsDir, 'office-coding-agent', 'office-excel');
    await fs.mkdir(path.join(pluginRoot, 'skills', 'excel'), { recursive: true });
    await fs.mkdir(path.join(pluginRoot, 'prompts'), { recursive: true });
    await fs.writeFile(
      path.join(pluginRoot, 'skills', 'excel', 'SKILL.md'),
      [
        '---',
        'name: excel',
        'description: Work with Excel workbooks',
        '---',
        '# Excel skill',
      ].join('\n')
    );
    await fs.writeFile(
      path.join(pluginRoot, 'prompts', 'cleanup.prompt.md'),
      [
        '---',
        'name: cleanup-workbook',
        'description: Clean up workbook data',
        '---',
        '# Cleanup prompt',
      ].join('\n')
    );

    const items = await getCliSlashItems({ installedPluginsDir });

    expect(items.skills).toEqual([
      {
        type: 'skill',
        name: 'excel',
        description: 'Work with Excel workbooks',
        plugin: 'office-excel@office-coding-agent',
      },
    ]);
    expect(items.prompts).toEqual([
      {
        type: 'prompt',
        name: 'cleanup-workbook',
        description: 'Clean up workbook data',
        plugin: 'office-excel@office-coding-agent',
      },
    ]);
  });

  it('returns empty lists when no Copilot CLI plugin directory exists', async () => {
    const items = await getCliSlashItems({
      installedPluginsDir: path.join(os.tmpdir(), 'office-cli-slash-items-missing'),
    });

    expect(items).toEqual({ skills: [], prompts: [] });
  });
});
