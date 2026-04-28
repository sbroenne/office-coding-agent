/**
 * start-dev-desktop.mjs — Starts the dev server, waits for it to be ready,
 * then sideloads the add-in into an Office desktop host.
 */

import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '..');

const APP_SCRIPT_NAMES = {
  excel: 'start:desktop:excel',
  powerpoint: 'start:desktop:ppt',
  ppt: 'start:desktop:ppt',
  word: 'start:desktop:word',
  outlook: 'start:desktop:outlook',
};

function getApp() {
  const appFlagIndex = process.argv.lastIndexOf('--app');
  const appValue = appFlagIndex >= 0 ? process.argv[appFlagIndex + 1] : undefined;
  return (appValue ?? 'excel').toLowerCase();
}

function runNpmScript(scriptName) {
  const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const result = spawnSync(npmCommand, ['run', scriptName], {
    cwd: PROJECT_ROOT,
    stdio: 'inherit',
  });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

const app = getApp();
const sideloadScriptName = APP_SCRIPT_NAMES[app];

if (!sideloadScriptName) {
  console.error(
    `Unsupported app "${app}". Use one of: excel, powerpoint, ppt, word, outlook.`
  );
  process.exit(1);
}

runNpmScript('dev:start');
runNpmScript(sideloadScriptName);
