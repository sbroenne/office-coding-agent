/**
 * kill-port.mjs — Kills any process listening on PORT before starting the dev server.
 *
 * Uses `netstat -ano` + `taskkill` — no PowerShell module required, works as a
 * standard user.  Waits up to 2 s for the port to be released before returning.
 */
import { execSync, spawnSync } from 'node:child_process';

const PORT = 3000;

function getPidsOnPort(port) {
  try {
    const out = execSync(`netstat -ano`, { encoding: 'utf-8', timeout: 5000 });
    const pids = new Set();
    for (const line of out.split('\n')) {
      // Match lines like "  TCP  0.0.0.0:3000  ...  LISTENING  12345"
      if (line.includes(`:${port}`) && line.includes('LISTENING')) {
        const pid = line.trim().split(/\s+/).at(-1);
        if (pid && /^\d+$/.test(pid) && pid !== '0') pids.add(Number(pid));
      }
    }
    return [...pids];
  } catch {
    return [];
  }
}

function isPortInUse(port) {
  return getPidsOnPort(port).length > 0;
}

// Kill all processes holding the port
const pids = getPidsOnPort(PORT);
for (const pid of pids) {
  spawnSync('taskkill', ['/pid', String(pid), '/f', '/t'], { timeout: 5000 });
  console.log(`  Killed PID ${pid} (was holding port ${PORT})`);
}

// Wait up to 2 s for the OS to release the port
if (pids.length > 0) {
  const deadline = Date.now() + 2000;
  while (Date.now() < deadline && isPortInUse(PORT)) {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 100);
  }
  if (!isPortInUse(PORT)) {
    console.log(`  Port ${PORT} is now free.`);
  } else {
    console.warn(`  Warning: port ${PORT} still in use after 2 s — proceeding anyway.`);
  }
}
