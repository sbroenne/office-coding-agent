/**
 * server.mjs — Express HTTPS server with Copilot WebSocket proxy + Vite dev middleware.
 *
 * Dev workflow:
 *   npm run dev          → starts this server (port 3000, HTTPS)
 *
 * The server:
 *  1. Serves the frontend via Vite (dev) or static dist (production)
 *  2. Proxies /api/copilot WebSocket to the @github/copilot-sdk
 *  3. Provides /api/fetch?url=... CORS proxy for the web_fetch tool
 *  4. Handles /api/upload-image for image attachments
 */

import express from 'express';
import cors from 'cors';
import https from 'node:https';
import http from 'node:http';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import { setupCopilotProxy, checkCopilotHealth } from './copilotProxy.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = 3000;
const isDev = process.env.NODE_ENV !== 'production';

/** Check that the port is available, exit early if it's in use. */
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

async function createServer() {
  console.log('\n  [server] Starting Copilot Office Add-in server...');
  await checkPort(PORT);

  const app = express();
  app.use(cors({ origin: '*' }));

  // ─── API Routes ──────────────────────────────────────────────────────────────
  const apiRouter = express.Router();
  apiRouter.use(express.json({ limit: '50mb' }));

  apiRouter.get('/hello', (_req, res) => {
    res.json({ message: 'Copilot proxy running', timestamp: new Date().toISOString() });
  });

  apiRouter.get('/ping', (_req, res) => {
    res.json({ ok: true });
  });

  // Copilot health check — reports whether any active session is connected
  apiRouter.get('/copilot-health', (_req, res) => {
    const health = checkCopilotHealth();
    res.json(health);
  });

  // CORS proxy for web_fetch tool (avoids CORS from Office add-in WebView)
  apiRouter.get('/fetch', (req, res) => {
    const url = req.query.url;
    if (!url) {
      res.status(400).json({ error: 'Missing url parameter' });
      return;
    }
    let parsedUrl;
    try {
      parsedUrl = new URL(url);
    } catch {
      res.status(400).json({ error: 'Invalid URL' });
      return;
    }
    // Only allow public HTTP/HTTPS — block SSRF to private/internal addresses
    if (parsedUrl.protocol !== 'http:' && parsedUrl.protocol !== 'https:') {
      res.status(400).json({ error: 'Only http and https URLs are allowed' });
      return;
    }
    const hostname = parsedUrl.hostname.toLowerCase();
    const BLOCKED = [
      /^localhost$/,
      /^127\./,
      /^0\.0\.0\.0$/,
      /^::1$/,
      /^10\./,
      /^172\.(1[6-9]|2\d|3[01])\./,
      /^192\.168\./,
      /^169\.254\./, // link-local / cloud metadata (AWS, Azure IMDS)
      /^fd[0-9a-f]{2}:/i, // IPv6 ULA
    ];
    if (BLOCKED.some(r => r.test(hostname))) {
      res.status(403).json({ error: 'Request to private or internal addresses is not allowed' });
      return;
    }
    try {
      const client = parsedUrl.protocol === 'https:' ? https : http;
      const options = {
        hostname: parsedUrl.hostname,
        path: parsedUrl.pathname + parsedUrl.search,
        headers: { 'User-Agent': 'OfficeAddinFetch/1.0' },
      };
      client
        .get(options, response => {
          let data = '';
          response.on('data', chunk => (data += chunk));
          response.on('end', () => {
            res.type('text/plain').send(data);
          });
        })
        .on('error', e => {
          res.status(500).json({ error: e.message });
        });
    } catch (e) {
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  // Image upload for multimodal prompts
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
      // path.basename prevents path traversal (e.g. name='../../etc/passwd')
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

  // ─── HTTPS server ───────────────────────────────────────────────────────────
  const devCerts = await import('office-addin-dev-certs');
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const httpsServer = https.createServer(httpsOptions, app);

  // ─── Copilot WebSocket Proxy (registered BEFORE Vite HMR) ────────────────────
  // Must come before createViteServer so the Copilot upgrade handler is the
  // first to receive WS upgrade events on /api/copilot — Vite's HMR handler
  // is registered afterwards and only consumes its own path.
  setupCopilotProxy(httpsServer);

  // ─── Frontend ────────────────────────────────────────────────────────────────
  if (isDev) {
    // Vite dev server in middleware mode.
    // Pass httpsServer via hmr.server so Vite attaches its HMR WebSocket to
    // our HTTPS server — without this, the Vite client can't upgrade to WS
    // and throws "WebSocket closed without opened."
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true, hmr: { server: httpsServer } },
      appType: 'custom',
    });
    app.use(vite.middlewares);

    // appType:'custom' disables Vite's HTML middleware — serve HTML manually
    // through Vite's transform pipeline so HMR client injection works
    const projectRoot = path.resolve(__dirname, '..');
    app.use(async (req, res, next) => {
      const isHtmlReq =
        req.url.endsWith('.html') || req.url === '/' || req.headers.accept?.includes('text/html');
      if (!isHtmlReq) return next();
      try {
        const htmlPath = path.join(projectRoot, 'taskpane.html');
        let html = fs.readFileSync(htmlPath, 'utf-8');
        html = await vite.transformIndexHtml(req.originalUrl, html);
        res.status(200).set({ 'Content-Type': 'text/html' }).end(html);
      } catch (e) {
        next(e);
      }
    });
  } else {
    // Serve static dist in production
    app.use(express.static(path.join(__dirname, '../dist')));
  }

  httpsServer.listen(PORT, () => {
    console.log(`\n  Copilot Office Add-in server running on https://localhost:${PORT}`);
    console.log(`  API: https://localhost:${PORT}/api\n`);
  });
}

createServer().catch(err => {
  console.error('Server startup error:', err);
  process.exit(1);
});
