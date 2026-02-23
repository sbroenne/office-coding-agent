import { spawn } from 'node:child_process';

function runNpmScript(scriptName) {
  return new Promise((resolve, reject) => {
    const child = spawn(`npm run ${scriptName}`, {
      shell: true,
      stdio: 'inherit',
      windowsHide: false,
    });

    child.on('exit', code => {
      if ((code ?? 1) === 0) {
        resolve();
      } else {
        reject(new Error(`npm run ${scriptName} failed with exit code ${String(code)}`));
      }
    });

    child.on('error', err => reject(err));
  });
}

async function main() {
  console.log('[stop:tray:desktop] Stopping Office sideload/debug session...');
  await runNpmScript('stop');

  console.log('[stop:tray:desktop] Stopping local server on port 3000...');
  await runNpmScript('dev:stop');

  console.log('[stop:tray:desktop] Done.');
}

main().catch(err => {
  console.error('[stop:tray:desktop] Failed:', err.message);
  process.exit(1);
});
