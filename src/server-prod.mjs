import express from 'express';
import cors from 'cors';
import https from 'node:https';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { setupCopilotProxy, checkCopilotHealth } from './copilotProxy.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = 3000;

async function checkPort(port) {
  return new Promise((resolve, reject) => {
    const tester = net
      .createServer()
      .once('error', () =>
        reject(
          new Error(
            `\n  ERROR: Port ${port} is already in use.\n  Stop the existing server and try again.\n`
          )
        )
      )
      .once('listening', () => tester.close(() => resolve()));
    tester.listen(port);
  });
}

export async function createServer() {
  await checkPort(PORT);

  const app = express();
  app.use(cors({ origin: '*' }));

  const apiRouter = express.Router();
  apiRouter.use(express.json({ limit: '50mb' }));

  apiRouter.get('/hello', (_req, res) => {
    res.json({ message: 'Copilot proxy running', timestamp: new Date().toISOString() });
  });

  apiRouter.get('/ping', (_req, res) => {
    res.json({ ok: true });
  });

  apiRouter.get('/env', (_req, res) => {
    res.json({
      cwd: process.cwd(),
      home: os.homedir(),
      platform: process.platform,
    });
  });

  apiRouter.get('/browse', async (req, res) => {
    try {
      const requestedPath = typeof req.query.path === 'string' ? req.query.path : process.cwd();
      const absolutePath = path.resolve(requestedPath);
      const entries = await fs.promises.readdir(absolutePath, { withFileTypes: true });
      const dirs = entries
        .filter(entry => entry.isDirectory())
        .map(entry => entry.name)
        .sort((a, b) => a.localeCompare(b));
      const parent = path.dirname(absolutePath);
      res.json({
        path: absolutePath,
        parent: parent === absolutePath ? null : parent,
        dirs,
      });
    } catch (error) {
      res.status(400).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/log', (req, res) => {
    const { level = 'error', tag = 'client', message, detail } = req.body || {};
    const prefix = `[${String(tag)}]`;
    if (level === 'error') {
      console.error(prefix, message, detail ?? '');
    } else {
      console.log(prefix, message, detail ?? '');
    }
    res.sendStatus(204);
  });

  apiRouter.get('/copilot-health', (_req, res) => {
    const health = checkCopilotHealth();
    res.json(health);
  });

  apiRouter.post('/upload-image', (req, res) => {
    try {
      const { dataUrl, name } = req.body;
      if (!dataUrl || !dataUrl.startsWith('data:image/')) {
        res.status(400).json({ error: 'Invalid image data' });
        return;
      }
      const matches = dataUrl.match(/^data:image\/([a-zA-Z+]+);base64,(.+)$/);
      if (!matches || matches.length !== 3) {
        res.status(400).json({ error: 'Invalid data URL format' });
        return;
      }
      const extension = matches[1] === 'svg+xml' ? 'svg' : matches[1];
      const base64Data = matches[2];
      const buffer = Buffer.from(base64Data, 'base64');
      const tempDir = path.join(os.tmpdir(), 'copilot-office-images');
      if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
      const filename = path.basename(name || `image-${Date.now()}.${extension}`);
      const filepath = path.join(tempDir, filename);
      fs.writeFileSync(filepath, buffer);
      res.json({ path: filepath, name: filename });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  app.use('/api', apiRouter);
  app.get('/ping', (_req, res) => res.json({ ok: true }));

  const devCerts = await import('office-addin-dev-certs');
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const httpsServer = https.createServer(httpsOptions, app);

  setupCopilotProxy(httpsServer);

  const distDir = path.resolve(__dirname, '..', 'dist');
  app.use(express.static(distDir));
  app.get('*path', (_req, res) => {
    res.sendFile(path.join(distDir, 'taskpane.html'));
  });

  await new Promise(resolve => {
    httpsServer.listen(PORT, () => {
      console.log(`\n  Copilot Office Add-in production server running on https://localhost:${PORT}`);
      console.log(`  API: https://localhost:${PORT}/api\n`);
      resolve(undefined);
    });
  });

  return httpsServer;
}

const isMainModule = process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1]);

if (isMainModule) {
  createServer().catch(err => {
    console.error('Server startup error:', err);
    process.exit(1);
  });
}
