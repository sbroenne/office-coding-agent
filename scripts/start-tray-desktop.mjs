import https from 'node:https';
import { spawn } from 'node:child_process';

const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const serverUrl = 'https://localhost:3000/api/ping';

function checkServerReady() {
  return new Promise(resolve => {
    const req = https.get(
      serverUrl,
      {
        rejectUnauthorized: false,
      },
      res => {
        res.resume();
        resolve(res.statusCode === 200);
      }
    );

    req.on('error', () => resolve(false));
    req.setTimeout(2000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

async function waitForServer(timeoutMs = 90_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await checkServerReady()) {
      return true;
    }
    await new Promise(resolve => setTimeout(resolve, 1000));
  }
  return false;
}

function startTrayDetached() {
  const tray = spawn(npmCommand, ['run', 'start:tray'], {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
  });
  tray.unref();
}

async function main() {
  const alreadyRunning = await checkServerReady();

  if (!alreadyRunning) {
    console.log('[start:tray:desktop] Starting tray app...');
    startTrayDetached();

    const ready = await waitForServer();
    if (!ready) {
      console.error('[start:tray:desktop] Tray server did not become ready at https://localhost:3000.');
      process.exit(1);
    }
  } else {
    console.log('[start:tray:desktop] Server already running on https://localhost:3000.');
  }

  console.log('[start:tray:desktop] Launching Excel sideload...');
  const sideload = spawn(npmCommand, ['run', 'start:desktop'], {
    stdio: 'inherit',
    windowsHide: false,
  });

  sideload.on('exit', code => {
    process.exit(code ?? 0);
  });

  sideload.on('error', err => {
    console.error('[start:tray:desktop] Failed to launch start:desktop:', err);
    process.exit(1);
  });
}

void main();
