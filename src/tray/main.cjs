const path = require('path');
const { spawn } = require('child_process');
const { app, Tray, Menu, shell, nativeImage } = require('electron');

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
  serverProcess = spawn(process.execPath, [serverPath], {
    cwd: path.resolve(__dirname, '../..'),
    env: {
      ...process.env,
      ELECTRON_RUN_AS_NODE: '1',
    },
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
