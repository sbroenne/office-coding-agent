import { spawn } from 'node:child_process';

export const OFFICE_CODING_AGENT_MARKETPLACE = {
  name: 'office-coding-agent',
  source: 'sbroenne/office-coding-agent-plugins',
};

export const REQUIRED_OFFICE_PLUGINS = [
  'office-excel',
  'office-powerpoint',
  'office-word',
  'office-outlook',
];

export const REQUIRED_OFFICE_PLUGIN_SPECS = REQUIRED_OFFICE_PLUGINS.map(
  name => `${name}@${OFFICE_CODING_AGENT_MARKETPLACE.name}`
);

const DEFAULT_TIMEOUT_MS = 60_000;

export function runCopilotPluginCommand(args, options = {}) {
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const command = options.command ?? 'copilot';

  return new Promise(resolve => {
    const child = spawn(command, ['plugin', ...args], {
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
      env: process.env,
    });

    let stdout = '';
    let stderr = '';
    let settled = false;
    const timer = setTimeout(() => {
      settled = true;
      child.kill();
      resolve({
        success: false,
        stdout,
        stderr,
        message: `copilot plugin ${args.join(' ')} timed out after ${timeoutMs}ms`,
      });
    }, timeoutMs);

    child.stdout?.on('data', chunk => {
      stdout += chunk.toString();
    });
    child.stderr?.on('data', chunk => {
      stderr += chunk.toString();
    });

    child.on('error', error => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({
        success: false,
        stdout,
        stderr,
        message: error.message,
      });
    });

    child.on('close', code => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      const message = (stdout || stderr).trim();
      resolve({
        success: code === 0,
        stdout,
        stderr,
        message,
        code,
      });
    });
  });
}

export function hasMarketplace(listOutput, marketplace = OFFICE_CODING_AGENT_MARKETPLACE) {
  return (
    listOutput.includes(marketplace.name) ||
    listOutput.includes(marketplace.source)
  );
}

export function installedPluginSpecs(listOutput) {
  const specs = new Set();
  for (const match of listOutput.matchAll(/([A-Za-z0-9._-]+@[A-Za-z0-9._-]+)/g)) {
    specs.add(match[1]);
  }
  return specs;
}

function logResult(logger, action, result) {
  const message = result.message?.trim();
  if (result.success) {
    logger.log(`  [plugin-bootstrap] ${action}`);
    if (message) logger.log(`  [plugin-bootstrap] ${message}`);
    return;
  }

  logger.warn(`  [plugin-bootstrap] ${action} failed: ${message || 'unknown error'}`);
}

export async function ensureOfficeCliPlugins(options = {}) {
  const run = options.runCommand ?? runCopilotPluginCommand;
  const logger = options.logger ?? console;
  const marketplace = options.marketplace ?? OFFICE_CODING_AGENT_MARKETPLACE;
  const pluginSpecs = options.pluginSpecs ?? REQUIRED_OFFICE_PLUGIN_SPECS;

  try {
    const marketplaceList = await run(['marketplace', 'list']);
    if (!marketplaceList.success) {
      logResult(logger, 'Marketplace list', marketplaceList);
    } else if (!hasMarketplace(marketplaceList.stdout, marketplace)) {
      logResult(
        logger,
        `Register marketplace ${marketplace.name}`,
        await run(['marketplace', 'add', marketplace.source])
      );
    }

    logResult(
      logger,
      `Update marketplace ${marketplace.name}`,
      await run(['marketplace', 'update', marketplace.name])
    );

    const pluginList = await run(['list']);
    const installed = pluginList.success ? installedPluginSpecs(pluginList.stdout) : new Set();
    if (!pluginList.success) {
      logResult(logger, 'Plugin list', pluginList);
    }

    for (const spec of pluginSpecs) {
      if (!installed.has(spec)) {
        logResult(logger, `Install plugin ${spec}`, await run(['install', spec]));
      }

      logResult(logger, `Update plugin ${spec}`, await run(['update', spec]));
    }

    logger.log('  [plugin-bootstrap] Office CLI plugin setup complete');
    return { success: true };
  } catch (error) {
    logger.warn(
      `  [plugin-bootstrap] Office CLI plugin setup failed (non-fatal): ${
        error instanceof Error ? error.message : String(error)
      }`
    );
    return { success: false, error };
  }
}
