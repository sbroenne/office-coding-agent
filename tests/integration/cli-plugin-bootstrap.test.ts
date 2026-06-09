import { describe, it, expect, vi } from 'vitest';
import {
  ensureOfficeCliPlugins,
  getInstalledOfficePluginDirectories,
  hasMarketplace,
  installedPluginSpecs,
  OFFICE_CODING_AGENT_MARKETPLACE,
  REQUIRED_OFFICE_PLUGINS,
  REQUIRED_OFFICE_PLUGIN_SPECS,
} from '@/../src/plugins/cliPluginBootstrap.mjs';

type CommandResult = {
  success: boolean;
  stdout: string;
  stderr?: string;
  message?: string;
};

const ok = (stdout = ''): CommandResult => ({ success: true, stdout, message: stdout });
const fail = (message: string): CommandResult => ({
  success: false,
  stdout: '',
  stderr: message,
  message,
});

describe('CLI plugin bootstrap', () => {
  it('detects the Office Coding Agent marketplace from CLI list output', () => {
    expect(
      hasMarketplace(
        'Registered marketplaces:\n  - office-coding-agent (GitHub: sbroenne/office-coding-agent-plugins)'
      )
    ).toBe(true);
    expect(hasMarketplace('Registered marketplaces:\n  - other (GitHub: owner/repo)')).toBe(false);
  });

  it('parses installed plugin specs from CLI list output', () => {
    expect(
      installedPluginSpecs(
        'Installed plugins:\n  - office-excel@office-coding-agent (v1.0.0)\n  - calendar@local (v1.0.0)'
      )
    ).toEqual(new Set(['office-excel@office-coding-agent', 'calendar@local']));
  });

  it('registers the marketplace, installs missing plugins, and updates required plugins', async () => {
    const runCommand = vi.fn(async (args: string[]) => {
      const key = args.join(' ');
      if (key === 'marketplace list') return ok('Registered marketplaces:\n');
      if (key === 'list') return ok('Installed plugins:\n  - office-excel@office-coding-agent (v1.0.0)');
      return ok(`${key} ok`);
    });
    const logger = { log: vi.fn(), warn: vi.fn() };

    await ensureOfficeCliPlugins({ runCommand, logger });

    expect(runCommand).toHaveBeenCalledWith(['marketplace', 'add', OFFICE_CODING_AGENT_MARKETPLACE.source]);
    expect(runCommand).toHaveBeenCalledWith(['marketplace', 'update', OFFICE_CODING_AGENT_MARKETPLACE.name]);
    expect(runCommand).not.toHaveBeenCalledWith(['install', 'office-excel@office-coding-agent']);
    for (const spec of REQUIRED_OFFICE_PLUGIN_SPECS.filter(spec => spec !== 'office-excel@office-coding-agent')) {
      expect(runCommand).toHaveBeenCalledWith(['install', spec]);
    }
    for (const spec of REQUIRED_OFFICE_PLUGIN_SPECS) {
      expect(runCommand).toHaveBeenCalledWith(['update', spec]);
    }
    expect(logger.warn).not.toHaveBeenCalled();
  });

  it('skips marketplace registration and plugin installs when already present', async () => {
    const runCommand = vi.fn(async (args: string[]) => {
      if (args.join(' ') === 'marketplace list') {
        return ok(`Registered marketplaces:\n  - ${OFFICE_CODING_AGENT_MARKETPLACE.name}`);
      }
      if (args.join(' ') === 'list') {
        return ok(REQUIRED_OFFICE_PLUGIN_SPECS.join('\n'));
      }
      return ok();
    });

    await ensureOfficeCliPlugins({ runCommand, logger: { log: vi.fn(), warn: vi.fn() } });

    expect(runCommand).not.toHaveBeenCalledWith(['marketplace', 'add', OFFICE_CODING_AGENT_MARKETPLACE.source]);
    for (const spec of REQUIRED_OFFICE_PLUGIN_SPECS) {
      expect(runCommand).not.toHaveBeenCalledWith(['install', spec]);
      expect(runCommand).toHaveBeenCalledWith(['update', spec]);
    }
  });

  it('logs CLI failures but keeps startup non-fatal', async () => {
    const runCommand = vi.fn(async (args: string[]) => {
      if (args.join(' ') === 'marketplace list') return fail('list failed');
      if (args.join(' ') === 'list') return fail('plugin list failed');
      return ok();
    });
    const logger = { log: vi.fn(), warn: vi.fn() };

    await expect(ensureOfficeCliPlugins({ runCommand, logger })).resolves.toEqual({
      success: true,
    });
    expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('Marketplace list failed'));
    expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('Plugin list failed'));
  });

  describe('getInstalledOfficePluginDirectories', () => {
    it('resolves existing Office plugin dirs under the marketplace install path', () => {
      const home = '/fake/home/.copilot';
      const present = new Set(
        REQUIRED_OFFICE_PLUGINS.map(
          name => `${home}/installed-plugins/${OFFICE_CODING_AGENT_MARKETPLACE.name}/${name}`
        )
      );
      const dirs = getInstalledOfficePluginDirectories({
        home,
        fileExists: (dir: string) => present.has(dir.replace(/\\/g, '/')),
      });

      expect(dirs).toHaveLength(REQUIRED_OFFICE_PLUGINS.length);
      for (const name of REQUIRED_OFFICE_PLUGINS) {
        expect(dirs.some(dir => dir.replace(/\\/g, '/').endsWith(`/${name}`))).toBe(true);
      }
    });

    it('skips plugin dirs that are not installed on disk', () => {
      const home = '/fake/home/.copilot';
      const installed = `${home}/installed-plugins/${OFFICE_CODING_AGENT_MARKETPLACE.name}/office-excel`;
      const dirs = getInstalledOfficePluginDirectories({
        home,
        fileExists: (dir: string) => dir.replace(/\\/g, '/') === installed,
      });

      expect(dirs.map(dir => dir.replace(/\\/g, '/'))).toEqual([installed]);
    });

    it('returns an empty list when no plugins are installed', () => {
      expect(
        getInstalledOfficePluginDirectories({ home: '/fake/home/.copilot', fileExists: () => false })
      ).toEqual([]);
    });
  });
});
