// @vitest-environment node

import { describe, expect, it, vi } from 'vitest';
import {
  autoSetupPlugins,
  HOST_PLUGINS,
  OCA_MARKETPLACE,
  OCA_MARKETPLACE_NAME,
} from '../../src/plugins/pluginAutoSetup.mjs';

type MockWithCalls = { mock: { calls: unknown[][] } };

function createLogger() {
  return {
    log: vi.fn(),
    warn: vi.fn(),
  };
}

function calledSpecs(fn: MockWithCalls) {
  return fn.mock.calls.map(call => call[0]);
}

function createManager(installed: string[]) {
  return {
    refresh: vi.fn(async () => undefined),
    list: vi.fn(() => installed.map(name => ({ name }))),
  };
}

describe('plugin auto setup', () => {
  it('registers the Office Coding Agent marketplace and installs all host plugins when missing', async () => {
    const manager = createManager([]);
    const addPluginMarketplaceFn = vi.fn(async () => ({ success: true }));
    const installPluginFn = vi.fn(async () => ({ success: true }));

    await autoSetupPlugins({
      getPluginManagerFn: async () => manager,
      listPluginMarketplacesFn: async () => ({ success: true, data: [] }),
      addPluginMarketplaceFn,
      installPluginFn,
      logger: createLogger(),
    });

    expect(addPluginMarketplaceFn).toHaveBeenCalledWith(OCA_MARKETPLACE);
    expect(manager.refresh).toHaveBeenCalledOnce();
    expect(installPluginFn).toHaveBeenCalledTimes(HOST_PLUGINS.length);
    expect(calledSpecs(installPluginFn)).toEqual(
      HOST_PLUGINS.map(name => `${name}@${OCA_MARKETPLACE_NAME}`)
    );
  });

  it('does not reinstall bundled host plugins that are already present', async () => {
    const manager = createManager(['office-excel', 'office-word']);
    const addPluginMarketplaceFn = vi.fn(async () => ({ success: true }));
    const installPluginFn = vi.fn(async () => ({ success: true }));

    await autoSetupPlugins({
      getPluginManagerFn: async () => manager,
      listPluginMarketplacesFn: async () => ({
        success: true,
        data: [{ name: OCA_MARKETPLACE_NAME }],
      }),
      addPluginMarketplaceFn,
      installPluginFn,
      logger: createLogger(),
    });

    expect(addPluginMarketplaceFn).not.toHaveBeenCalled();
    expect(calledSpecs(installPluginFn)).toEqual([
      `office-powerpoint@${OCA_MARKETPLACE_NAME}`,
      `office-outlook@${OCA_MARKETPLACE_NAME}`,
    ]);
  });

  it('keeps startup non-fatal when a bundled plugin install fails', async () => {
    const logger = createLogger();
    const installPluginFn = vi
      .fn()
      .mockResolvedValueOnce({ success: false, error: 'network unavailable' })
      .mockResolvedValue({ success: true });

    await expect(
      autoSetupPlugins({
        getPluginManagerFn: async () => createManager([]),
        listPluginMarketplacesFn: async () => ({
          success: true,
          data: [{ name: OCA_MARKETPLACE_NAME }],
        }),
        addPluginMarketplaceFn: async () => ({ success: true }),
        installPluginFn,
        logger,
      })
    ).resolves.toBeUndefined();

    expect(installPluginFn).toHaveBeenCalledTimes(HOST_PLUGINS.length);
    expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('office-excel install failed'));
  });

  it('keeps server startup non-fatal if plugin setup throws', async () => {
    const logger = createLogger();

    await expect(
      autoSetupPlugins({
        getPluginManagerFn: async () => {
          throw new Error('config unreadable');
        },
        logger,
      })
    ).resolves.toBeUndefined();

    expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('config unreadable'));
  });
});
