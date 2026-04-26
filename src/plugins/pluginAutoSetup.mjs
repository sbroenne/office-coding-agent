import {
  addPluginMarketplace,
  installPlugin,
  listPluginMarketplaces,
} from './pluginCli.mjs';
import { getPluginManager } from './pluginManager.mjs';

export const OCA_MARKETPLACE = 'sbroenne/office-coding-agent-plugins';
export const OCA_MARKETPLACE_NAME = 'office-coding-agent';
export const HOST_PLUGINS = ['office-excel', 'office-powerpoint', 'office-word', 'office-outlook'];

export async function autoSetupPlugins(deps = {}) {
  const {
    getPluginManagerFn = getPluginManager,
    listPluginMarketplacesFn = listPluginMarketplaces,
    addPluginMarketplaceFn = addPluginMarketplace,
    installPluginFn = installPlugin,
    logger = console,
  } = deps;

  try {
    const manager = await getPluginManagerFn();
    const marketplaceResult = await listPluginMarketplacesFn();
    const marketplaces = marketplaceResult.data ?? [];

    const hasMarketplace = marketplaces.some(m => m.name === OCA_MARKETPLACE_NAME);
    if (!hasMarketplace) {
      logger.log(`  [auto-setup] Registering marketplace: ${OCA_MARKETPLACE}`);
      const result = await addPluginMarketplaceFn(OCA_MARKETPLACE);
      if (result.success) {
        logger.log('  [auto-setup] Marketplace registered');
      } else {
        logger.warn(`  [auto-setup] Marketplace registration failed: ${result.error}`);
      }
    }

    await manager.refresh();
    const installed = manager.list();
    for (const pluginName of HOST_PLUGINS) {
      const isInstalled = installed.some(p => p.name === pluginName);
      if (!isInstalled) {
        logger.log(`  [auto-setup] Installing plugin: ${pluginName}`);
        const result = await installPluginFn(`${pluginName}@${OCA_MARKETPLACE_NAME}`);
        if (result.success) {
          logger.log(`  [auto-setup] ${pluginName} installed`);
        } else {
          logger.warn(`  [auto-setup] ${pluginName} install failed: ${result.error}`);
        }
      }
    }

    logger.log('  [auto-setup] Plugin setup complete');
  } catch (err) {
    logger.warn(`  [auto-setup] Auto-setup failed (non-fatal): ${err.message}`);
  }
}
