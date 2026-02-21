/**
 * start-dev-server.mjs — Launches `npm run dev` in an external CMD window
 * and waits until the server is listening on the port.
 *
 * Replicates the office-addin-debugging pattern:
 *   1. Kill any stale process on the port
 *   2. Spawn `npm run dev` in a detached CMD window (visible)
 *   3. Save the CMD process PID to a JSON file in TEMP
 *   4. Poll the port until the server is listening (up to 30s)
 *
 * The companion `kill-port.mjs` / `npm run stop` reads the PID file
 * and uses `taskkill /t` to tear down the whole process tree.
 */

import { spawn, execSync } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '..');
const PORT = 3000;
const PID_FILE = path.join(os.tmpdir(), 'office-coding-agent-dev-server.json');
const MAX_RETRIES = 30;
const RETRY_DELAY_MS = 1000;

// ── Helpers ────────────────────────────────────────────────────────────────────

/** Kill any process occupying the port (Windows-only). */
function killPortHolder(port) {
  try {
    execSync(
      `powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort ${port} -ErrorAction SilentlyContinue | Where-Object { $_.OwningProcess -ne 0 } | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { Stop-Process -Id $_ -Force }"`,
      { stdio: 'ignore', timeout: 5000 }
    );
  } catch {
    // no process on port — ok
  }
}

/** Kill a process tree by PID using taskkill (Windows-only). */
function killProcessTree(pid) {
  try {
    execSync(`taskkill /pid ${pid} /f /t`, { stdio: 'ignore', timeout: 5000 });
  } catch {
    // process already gone — ok
  }
}

/** Check if any process is listening on the port. */
function isPortListening(port) {
  return new Promise(resolve => {
    const socket = new net.Socket();
    socket.setTimeout(500);
    socket.once('connect', () => {
      socket.destroy();
      resolve(true);
    });
    socket.once('timeout', () => {
      socket.destroy();
      resolve(false);
    });
    socket.once('error', () => {
      resolve(false);
    });
    socket.connect(port, '127.0.0.1');
  });
}

/** Wait until a process is listening on the port. */
async function waitForServer(port, retries, delayMs) {
  for (let i = 0; i < retries; i++) {
    if (await isPortListening(port)) return true;
    await new Promise(r => setTimeout(r, delayMs));
  }
  return false;
}

/** Get the PID of the process listening on the port (Windows-only). */
function getProcessOnPort(port) {
  try {
    const output = execSync(
      `powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort ${port} -ErrorAction SilentlyContinue | Where-Object { $_.OwningProcess -ne 0 } | Select-Object -ExpandProperty OwningProcess -Unique -First 1"`,
      { encoding: 'utf-8', timeout: 5000 }
    ).trim();
    const pid = parseInt(output, 10);
    return Number.isFinite(pid) ? pid : null;
  } catch {
    return null;
  }
}

/** Walk up the process tree to find the top-level ancestor (the CMD window). */
function getRootAncestor(pid) {
  let current = pid;
  for (let i = 0; i < 10; i++) {
    try {
      const output = execSync(
        `powershell -NoProfile -Command "(Get-CimInstance Win32_Process -Filter 'ProcessId=${current}').ParentProcessId"`,
        { encoding: 'utf-8', timeout: 5000 }
      ).trim();
      const parentPid = parseInt(output, 10);
      if (!Number.isFinite(parentPid) || parentPid === 0) break;
      // Check if the parent is a cmd.exe or node.exe we spawned
      const parentName = execSync(
        `powershell -NoProfile -Command "(Get-Process -Id ${parentPid} -ErrorAction SilentlyContinue).ProcessName"`,
        { encoding: 'utf-8', timeout: 5000 }
      )
        .trim()
        .toLowerCase();
      if (parentName === 'cmd' || parentName === 'node') {
        current = parentPid;
      } else {
        break; // parent is explorer, pwsh, etc. — stop here
      }
    } catch {
      break;
    }
  }
  return current;
}

/** Save the dev server PID to a JSON file. */
function savePid(pid) {
  fs.writeFileSync(PID_FILE, JSON.stringify({ devServer: { processId: pid } }));
}

/** Read a previously saved PID. Returns null if not found. */
function readPid() {
  try {
    const data = JSON.parse(fs.readFileSync(PID_FILE, 'utf-8'));
    return data?.devServer?.processId ?? null;
  } catch {
    return null;
  }
}

// ── Main ───────────────────────────────────────────────────────────────────────

// Fast path: if server is already listening, reuse it — don't kill/restart.
// This avoids VS Code's "terminate existing instances?" dialog on repeated F5.
if (await isPortListening(PORT)) {
  console.log(`  Dev server is already running on https://localhost:${PORT}`);
  process.exit(0);
}

// 1. Kill any stale dev server (by saved PID, then by port)
const stalePid = readPid();
if (stalePid) {
  console.log(`  Killing stale dev server (PID ${stalePid})...`);
  killProcessTree(stalePid);
  await new Promise(r => setTimeout(r, 1000));
}
// Always clean up port as safety net (covers orphaned processes)
if (await isPortListening(PORT)) {
  console.log(`  Port ${PORT} still in use — killing port holder...`);
  killPortHolder(PORT);
  await new Promise(r => setTimeout(r, 1500));
}

// 2. Spawn `npm run dev` in an external CMD window using `cmd /c start`
//    `start` always opens a new visible console window on Windows.
const cmdArgs = `cd /d ${PROJECT_ROOT} && npm run dev`;
const child = spawn('cmd', ['/c', 'start', 'Dev Server', 'cmd', '/c', cmdArgs], {
  cwd: PROJECT_ROOT,
  detached: true,
  stdio: 'ignore',
});
child.unref();

console.log(`  Dev server launching in external window...`);

// 3. Wait for the server to start listening, then save the root CMD window PID
const ready = await waitForServer(PORT, MAX_RETRIES, RETRY_DELAY_MS);
if (ready) {
  // Find the process on the port, then walk up to the CMD window root
  const serverPid = getProcessOnPort(PORT);
  const rootPid = serverPid ? getRootAncestor(serverPid) : null;
  if (rootPid) {
    savePid(rootPid);
    console.log(`  Dev server is running on https://localhost:${PORT} (tree root PID ${rootPid})`);
  } else {
    console.log(`  Dev server is running on https://localhost:${PORT}`);
  }
} else {
  console.error(`  ERROR: Dev server did not start within ${MAX_RETRIES}s.`);
  process.exit(1);
}
