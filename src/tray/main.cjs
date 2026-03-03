const path = require('path');
const fs = require('fs');
const { spawn, execFileSync } = require('child_process');
const { app, Tray, Menu, shell, nativeImage } = require('electron');

/**
 * Find a real Node.js executable on PATH, skipping the Electron binary.
 * Returns the found path, or null if not found.
 */
function findRealNode() {
  // Already explicitly provided and valid
  const configured = process.env.ORIGINAL_NODE_EXE;
  if (configured && fs.existsSync(configured)) return configured;

  try {
    const cmd = process.platform === 'win32' ? 'where' : 'which';
    const result = execFileSync(cmd, ['node'], { encoding: 'utf8', timeout: 3000 });
    const candidates = result
      .trim()
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean);

    for (const candidate of candidates) {
      // Skip Electron's bundled Node impersonation
      if (!candidate.toLowerCase().includes('electron')) {
        return candidate;
      }
    }
  } catch {
    // PATH search failed — fall through
  }
  return null;
}

let tray = null;
let serverProcess = null;
let serverStatus = 'stopped';
let lastServerError = null;

const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) {
  app.quit();
}

app.on('second-instance', () => {
  if (tray) {
    tray.popUpContextMenu();
  }
});

function getIconPath() {
  return path.resolve(__dirname, '../../assets/icon-32.png');
}

function startServer() {
  if (serverProcess) return;

  serverStatus = 'starting';
  lastServerError = null;
  updateMenu();

  const serverPath = path.resolve(__dirname, '../server-prod.mjs');
  const preferredNode = process.env.ORIGINAL_NODE_EXE;
  const useRealNode = Boolean(preferredNode && fs.existsSync(preferredNode));
  const runtime = useRealNode ? preferredNode : process.execPath;

  const env = {
    ...process.env,
  };

  if (useRealNode) {
    delete env.ELECTRON_RUN_AS_NODE;
  } else {
    env.ELECTRON_RUN_AS_NODE = '1';
  }

  serverProcess = spawn(runtime, [serverPath], {
    cwd: path.resolve(__dirname, '../..'),
    env,
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  serverProcess.stdout.on('data', data => {
    const text = data.toString();
    process.stdout.write(`[tray-server] ${text}`);
    if (text.includes('production server running on https://localhost:3000')) {
      serverStatus = 'running';
      updateMenu();
    }
  });

  serverProcess.stderr.on('data', data => {
    const text = data.toString();
    process.stderr.write(`[tray-server] ${text}`);
    const trimmed = text.trim();
    if (trimmed.length > 0) {
      lastServerError = trimmed;
      if (serverStatus !== 'running') {
        serverStatus = 'error';
      }
      updateMenu();
    }
  });

  serverProcess.on('exit', code => {
    console.log(`[tray-server] exited with code ${String(code)}`);
    if (code !== 0 && code !== null) {
      serverStatus = 'error';
      if (!lastServerError) {
        lastServerError = `Server exited with code ${String(code)}`;
      }
    } else {
      serverStatus = 'stopped';
    }
    serverProcess = null;
    updateMenu();
  });
}

function stopServer() {
  if (!serverProcess) return;
  serverProcess.kill();
  serverProcess = null;
  serverStatus = 'stopped';
}

function updateMenu() {
  const statusLabel =
    serverStatus === 'running'
      ? 'Server: Running'
      : serverStatus === 'starting'
        ? 'Server: Starting'
        : serverStatus === 'error'
          ? 'Server: Error'
          : 'Server: Stopped';

  const tooltip =
    serverStatus === 'running'
      ? 'Office Coding Agent (running)'
      : serverStatus === 'starting'
        ? 'Office Coding Agent (starting)'
        : serverStatus === 'error'
          ? 'Office Coding Agent (error)'
          : 'Office Coding Agent (stopped)';

  const menu = Menu.buildFromTemplate([
    {
      label: statusLabel,
      enabled: false,
    },
    {
      label: lastServerError ? `Last error: ${lastServerError}` : 'Last error: none',
      enabled: false,
    },
    {
      label: 'Open API Health',
      click: () => shell.openExternal('https://localhost:3000/api/ping'),
    },
    { type: 'separator' },
    {
      label: 'Restart Server',
      enabled: serverStatus !== 'starting',
      click: () => {
        stopServer();
        startServer();
        updateMenu();
      },
    },
    {
      label: 'Quit',
      click: () => {
        app.quit();
      },
    },
  ]);
  tray.setContextMenu(menu);
  tray.setToolTip(tooltip);
}

app.whenReady().then(() => {
  // Ensure the Copilot SDK spawns the CLI with real Node, not Electron.
  // The SDK uses process.execPath (= Electron exe in tray mode) to run the
  // bundled CLI JS file, which causes Electron to inject extra argv entries
  // that the CLI rejects as "too many arguments". Setting ORIGINAL_NODE_EXE
  // causes startServer() to use real Node as the server runtime, making
  // process.execPath correct for any child process the SDK spawns.
  const realNode = findRealNode();
  if (realNode) {
    process.env.ORIGINAL_NODE_EXE = realNode;
    console.log('[tray] Using real Node for server:', realNode);
  } else {
    console.warn('[tray] Could not find real Node on PATH; Copilot CLI may fail to start.');
  }

  const icon = nativeImage.createFromPath(getIconPath());
  tray = new Tray(icon);
  startServer();
  updateMenu();

  tray.on('click', () => {
    tray.popUpContextMenu();
  });
});

app.on('window-all-closed', e => {
  e.preventDefault();
});

app.on('before-quit', () => {
  stopServer();
});
