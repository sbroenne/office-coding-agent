/**
 * Node.js-only helpers for the Word E2E test runner.
 * NOT bundled by Vite — used only by runner.test.ts (Mocha/Node).
 */

import * as childProcess from 'child_process';

/* global process */

/**
 * Close the Word desktop application.
 */
export async function closeDesktopApplication(): Promise<boolean> {
  try {
    if (process.platform === 'win32') {
      return await executeCommandLine('tskill WINWORD');
    }
    return false;
  } catch {
    throw new Error('Unable to kill WINWORD process.');
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
