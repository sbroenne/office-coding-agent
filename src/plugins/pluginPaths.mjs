import { homedir } from 'node:os';
import { join } from 'node:path';

const APP_DATA_OVERRIDE = 'OFFICE_CODING_AGENT_APP_DATA';

export function getAppDataDir() {
  if (process.env[APP_DATA_OVERRIDE]) return process.env[APP_DATA_OVERRIDE];

  if (process.platform === 'win32') {
    const base = process.env.APPDATA || join(homedir(), 'AppData', 'Roaming');
    return join(base, 'Office Coding Agent');
  }

  if (process.platform === 'darwin') {
    return join(homedir(), 'Library', 'Application Support', 'Office Coding Agent');
  }

  return join(process.env.XDG_CONFIG_HOME || join(homedir(), '.config'), 'office-coding-agent');
}

export function getCliHomeDir() {
  return join(getAppDataDir(), 'cli-home');
}

export function getCopilotHomeDir() {
  return join(getCliHomeDir(), '.copilot');
}

export function getCopilotConfigPath() {
  return join(getCopilotHomeDir(), 'config.json');
}

export function getMarketplaceCacheDir() {
  return join(getCopilotHomeDir(), 'marketplace-cache');
}
