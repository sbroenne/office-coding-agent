const path = require('path');
const { spawn } = require('child_process');
const { app, Tray, Menu, shell, nativeImage } = require('electron');

let tray = null;
let serverProcess = null;

function getIconPath() {
  return path.resolve(__dirname, '../../assets/icon-32.png');
}

function startServer() {
  if (serverProcess) return;

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
    process.stdout.write(`[tray-server] ${data.toString()}`);
  });

  serverProcess.stderr.on('data', data => {
    process.stderr.write(`[tray-server] ${data.toString()}`);
  });

  serverProcess.on('exit', code => {
    console.log(`[tray-server] exited with code ${String(code)}`);
    serverProcess = null;
    updateMenu();
  });
}

function stopServer() {
  if (!serverProcess) return;
  serverProcess.kill();
  serverProcess = null;
}

function updateMenu() {
  const running = !!serverProcess;
  const menu = Menu.buildFromTemplate([
    {
      label: running ? 'Server: Running' : 'Server: Stopped',
      enabled: false,
    },
    {
      label: 'Open API Health',
      click: () => shell.openExternal('https://localhost:3000/api/ping'),
    },
    { type: 'separator' },
    {
      label: 'Restart Server',
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
  tray.setToolTip(running ? 'Office Coding Agent (running)' : 'Office Coding Agent (stopped)');
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
