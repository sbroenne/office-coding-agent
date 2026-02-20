/* eslint-disable @typescript-eslint/no-require-imports */
/**
 * server.js — Express HTTPS server with Copilot WebSocket proxy + webpack dev middleware.
 *
 * Dev workflow:
 *   npm run server        → starts this server (port 3000, HTTPS, replaces webpack-dev-server)
 *   npm run server:build  → starts this server alongside webpack build watcher
 *
 * The server:
 *  1. Serves static webpack output (or uses webpack-dev-middleware in dev mode)
 *  2. Proxies /api/copilot WebSocket to the @github/copilot CLI
 *  3. Provides /api/fetch?url=... CORS proxy for the web_fetch tool
 *  4. Handles /api/upload-image for image attachments
 */

const express = require('express');
const cors = require('cors');
const https = require('https');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { setupCopilotProxy } = require('./copilotProxy');

const PORT = 3000;
const isDev = process.env.NODE_ENV !== 'production';

async function createServer() {
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

  // CORS proxy for web_fetch tool (avoids CORS from Office add-in WebView)
  apiRouter.get('/fetch', (req, res) => {
    const url = req.query.url;
    if (!url) {
      res.status(400).json({ error: 'Missing url parameter' });
      return;
    }
    try {
      const http = require('http');
      const parsedUrl = new URL(url);
      const client = parsedUrl.protocol === 'https:' ? https : http;
      const options = {
        hostname: parsedUrl.hostname,
        path: parsedUrl.pathname + parsedUrl.search,
        headers: { 'User-Agent': 'OfficeAddinFetch/1.0' },
      };
      client.get(options, response => {
        let data = '';
        response.on('data', chunk => (data += chunk));
        response.on('end', () => {
          res.type('text/plain').send(data);
        });
      }).on('error', e => {
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
      const filename = name || `image-${Date.now()}.${extension}`;
      const filepath = path.join(tempDir, filename);
      fs.writeFileSync(filepath, buffer);
      res.json({ path: filepath, name: filename });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  app.use('/api', apiRouter);
  app.get('/ping', (_req, res) => res.json({ ok: true }));

  // ─── Frontend ────────────────────────────────────────────────────────────────
  if (isDev) {
    // Use webpack-dev-middleware in dev mode for HMR
    const webpack = require('webpack');
    const webpackDevMiddleware = require('webpack-dev-middleware');
    const webpackHotMiddleware = require('webpack-hot-middleware');
    const getWebpackConfig = require('../webpack.config.js');
    const config = await getWebpackConfig({}, { mode: 'development' });

    // Enable HMR entry
    if (!Array.isArray(config.entry.taskpane)) {
      config.entry.taskpane = ['webpack-hot-middleware/client?reload=true', config.entry.taskpane];
    }
    // Add HMR plugin
    if (!config.plugins) config.plugins = [];
    config.plugins.push(new webpack.HotModuleReplacementPlugin());

    const compiler = webpack(config);
    app.use(webpackDevMiddleware(compiler, { publicPath: '/' }));
    app.use(webpackHotMiddleware(compiler));
  } else {
    // Serve static dist in production
    app.use(express.static(path.join(__dirname, '../dist')));
  }

  // ─── HTTPS + WebSocket Proxy ──────────────────────────────────────────────
  const devCerts = require('office-addin-dev-certs');
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const httpsServer = https.createServer(httpsOptions, app);

  setupCopilotProxy(httpsServer);

  httpsServer.listen(PORT, () => {
    // eslint-disable-next-line no-console
    console.log(`\n  Copilot Office Add-in server running on https://localhost:${PORT}`);
    // eslint-disable-next-line no-console
    console.log(`  API: https://localhost:${PORT}/api\n`);
  });
}

createServer().catch(err => {
  // eslint-disable-next-line no-console
  console.error('Server startup error:', err);
  process.exit(1);
});
