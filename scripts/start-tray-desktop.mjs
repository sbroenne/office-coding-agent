import https from 'node:https';
import { spawn } from 'node:child_process';
const serverUrl = 'https://localhost:3000/api/ping';

function resolveAppArg() {
  const arg = process.argv.find(v => v.startsWith('--app='));
  const value = arg?.split('=')[1]?.toLowerCase() ?? 'excel';
  if (value === 'excel' || value === 'powerpoint' || value === 'word') return value;
  return 'excel';
}

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
  const tray = spawn('npm run start:tray', {
    shell: true,
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
  });
  tray.unref();
}

async function main() {
  const app = resolveAppArg();
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

  console.log(`[start:tray:desktop] Launching ${app} sideload...`);
  const scriptName =
    app === 'powerpoint' ? 'start:desktop:ppt' : app === 'word' ? 'start:desktop:word' : 'start:desktop:excel';

  const sideload = spawn(`npm run ${scriptName}`, {
    shell: true,
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
