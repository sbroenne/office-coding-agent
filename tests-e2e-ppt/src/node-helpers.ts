/**
 * Node.js-only helpers for the PowerPoint E2E test runner.
 * NOT bundled by Vite — used only by runner.test.ts (Mocha/Node).
 */

import * as childProcess from 'child_process';

/* global process */

/**
 * Close the PowerPoint desktop application.
 */
export async function closeDesktopApplication(): Promise<boolean> {
  try {
    if (process.platform === 'win32') {
      return await executeCommandLine('tskill POWERPNT');
    }
    return false;
  } catch {
    throw new Error('Unable to kill POWERPNT process.');
  }
}

/**
 * Execute a command line command.
 */
function executeCommandLine(cmdLine: string): Promise<boolean> {
  return new Promise(resolve => {
    childProcess.exec(cmdLine, error => {
      resolve(!error);
    });
  });
}
