/**
 * Node.js-only helpers for the Word E2E test runner.
 * NOT bundled by Vite — used only by runner.test.ts (Mocha/Node).
 */

import * as childProcess from 'child_process';

/* global process */

/**
 * Open the test pane only in the test-owned Word window on Windows.
 */
export async function openTestTaskpane(): Promise<void> {
  if (process.platform !== 'win32') return;
  const script = `
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
for ($attempt = 0; $attempt -lt 20; $attempt++) {
  $word = Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle.StartsWith('Word add-in e2e30000-ae5f-4000-a000-000000000003') } |
    Select-Object -First 1
  if ($word) {
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($word.MainWindowHandle)
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, 'Run Word E2E')
    $button = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
    if ($button) {
      $invoke = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
      $invoke.Invoke()
      exit 0
    }
  }
  Start-Sleep -Seconds 1
}
throw 'Could not open Run Word E2E in the test-owned Word window.'
`;
  await new Promise<void>((resolve, reject) => {
    childProcess.execFile(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-Command', script],
      { timeout: 30000 },
      (error, _stdout, stderr) => {
        if (error) reject(new Error(`Could not open Word test pane: ${stderr || error.message}`));
        else resolve();
      }
    );
  });
}
