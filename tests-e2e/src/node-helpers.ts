/**
 * Node.js-only helpers for the E2E test runner.
 * NOT bundled by Vite — used only by runner.test.ts (Mocha/Node).
 */

import * as childProcess from 'child_process';

/* global process */

export async function maximizeTestWindow(manifestId: string): Promise<void> {
  if (process.platform !== 'win32') return;
  if (!/^[0-9a-f-]{36}$/i.test(manifestId)) throw new Error('Invalid test manifest ID.');

  const script = `
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class TestExcelWindow { [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hwnd, int command); }'
    for ($i = 0; $i -lt 20; $i++) {
      $window = Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -like '*${manifestId}*' -and $_.MainWindowHandle -ne 0 } |
        Select-Object -First 1
      if ($window) {
        if (-not [TestExcelWindow]::ShowWindowAsync($window.MainWindowHandle, 3)) {
          throw 'Unable to maximize the Excel test workbook.'
        }
        exit 0
      }
      Start-Sleep -Seconds 1
    }
    throw 'The Excel test workbook window did not appear.'
  `;
  await new Promise<void>((resolve, reject) => {
    childProcess.execFile(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-Command', script],
      { timeout: 30_000 },
      error => (error ? reject(error) : resolve())
    );
  });
}

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
