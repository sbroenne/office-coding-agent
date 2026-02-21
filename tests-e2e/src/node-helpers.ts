/**
 * Node.js-only helpers for the E2E test runner.
 * NOT bundled by Vite — used only by runner.test.ts (Mocha/Node).
 */

import * as childProcess from 'child_process';

/* global process */

/**
 * Close the Excel desktop application.
 */
export async function closeDesktopApplication(): Promise<boolean> {
  try {
    if (process.platform === 'win32') {
      return await executeCommandLine('tskill Excel');
    }
    return false;
  } catch {
    throw new Error('Unable to kill Excel process.');
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
