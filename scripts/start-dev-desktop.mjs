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
  const command =
    process.platform === 'win32' ? (process.env.ComSpec ?? 'cmd.exe') : 'npm';
  const args =
    process.platform === 'win32' ? ['/d', '/c', 'npm', 'run', scriptName] : ['run', scriptName];
  const result = spawnSync(command, args, {
    cwd: PROJECT_ROOT,
    encoding: 'utf8',
  });

  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);

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
