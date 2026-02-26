/**
 * Node.js-only helpers for the Outlook E2E test runner.
 * NOT bundled by Vite — used only by runner.test.ts (Mocha/Node).
 */

import * as childProcess from 'child_process';

/* global process */

/**
 * Close the Outlook desktop application.
 */
export async function closeDesktopApplication(): Promise<boolean> {
  try {
    if (process.platform === 'win32') {
      return await executeCommandLine('tskill OUTLOOK');
    }
    return false;
  } catch {
    throw new Error('Unable to kill OUTLOOK process.');
  }
}

/**
 * Open a new compose window in Outlook Desktop.
 * This triggers the OnNewMessageCompose LaunchEvent which auto-runs the test taskpane.
 */
export async function openComposeWindow(): Promise<void> {
  if (process.platform !== 'win32') return;

  // Wait for Outlook to be fully ready before opening compose
  await sleep(5000);

  // Open new mail compose window via outlook.exe /c ipm.note
  await executeCommandLine('start "" "OUTLOOK.EXE" /c ipm.note');

  // Give the compose window time to open and the LaunchEvent to fire
  await sleep(3000);
}

/**
 * Sleep for a given number of milliseconds.
 */
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
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
