/**
 * server.mjs — Express HTTPS server with Copilot WebSocket proxy + Vite dev middleware.
 *
 * Dev workflow:
 *   npm run dev          → starts this server (port 3000, HTTPS)
 *
 * The server:
 *  1. Serves the frontend via Vite (dev) or static dist (production)
 *  2. Proxies /api/copilot WebSocket to the @github/copilot-sdk
 *  3. Handles /api/upload-image for image attachments
 */

import express from 'express';
import cors from 'cors';
import https from 'node:https';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import { setupCopilotProxy, checkCopilotHealth } from './copilotProxy.mjs';
import {
  addPluginMarketplace,
  browsePluginMarketplace,
  installPlugin,
  listPluginMarketplaces,
  removePluginMarketplace,
  uninstallPlugin,
  updatePlugin,
  updatePluginMarketplace,
} from './plugins/pluginCli.mjs';
import { autoSetupPlugins, OCA_MARKETPLACE_NAME } from './plugins/pluginAutoSetup.mjs';
import { getPluginManager } from './plugins/pluginManager.mjs';
import {
  getBrowseRoots,
  isAllowedOrigin,
  isPathWithinRoot,
  isTrustedRequestOrigin,
  resolveBrowsePath,
} from './serverSecurity.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = 3000;
const isDev = process.env.NODE_ENV !== 'production';

const BUNDLED_MCP_SERVERS = [
  {
    name: 'workiq',
    description: 'Microsoft 365 Copilot — emails, meetings, documents, Teams',
    transport: 'stdio',
    command: 'npx',
    args: ['-y', '@microsoft/workiq', 'mcp'],
  },
  {
    name: 'powerbi',
    description: 'Power BI — query semantic models, generate DAX, explore data',
    transport: 'http',
    url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi',
  },
];

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
  const browseRoots = await getBrowseRoots();

  app.use('/api', (req, res, next) => {
    const origin = req.headers.origin;
    if (origin && !isAllowedOrigin(origin)) {
      res.status(403).json({ error: 'Origin is not allowed.' });
      return;
    }
    next();
  });
  app.use(
    '/api',
    cors({
      origin(origin, callback) {
        if (!origin || isAllowedOrigin(origin)) {
          callback(null, true);
          return;
        }
        callback(null, false);
      },
    })
  );

  // ─── API Routes ──────────────────────────────────────────────────────────────
  const apiRouter = express.Router();
  apiRouter.use(express.json({ limit: '50mb' }));

  const requireTrustedLocalAccess = (req, res, next) => {
    const remoteAddress = req.socket?.remoteAddress;
    if (isTrustedRequestOrigin(req.headers.origin, remoteAddress)) {
      next();
      return;
    }

    res.status(403).json({ error: 'This endpoint is only available to the local add-in.' });
  };

  apiRouter.get('/hello', (_req, res) => {
    res.json({ message: 'Copilot proxy running', timestamp: new Date().toISOString() });
  });

  apiRouter.get('/ping', (_req, res) => {
    res.json({ ok: true });
  });

  apiRouter.get('/env', requireTrustedLocalAccess, (_req, res) => {
    res.json({
      platform: process.platform,
      nodeEnv: process.env.NODE_ENV ?? 'development',
      browseRestricted: true,
    });
  });

  apiRouter.get('/browse', requireTrustedLocalAccess, async (req, res) => {
    try {
      const requestedPath = typeof req.query.path === 'string' ? req.query.path : undefined;
      const absolutePath = await resolveBrowsePath(requestedPath, browseRoots);
      const entries = await fs.promises.readdir(absolutePath, { withFileTypes: true });
      const dirs = entries
        .filter(entry => entry.isDirectory())
        .map(entry => entry.name)
        .sort((a, b) => a.localeCompare(b));
      const parent = path.dirname(absolutePath);
      const parentAllowed = parent !== absolutePath && browseRoots.some(root => isPathWithinRoot(root, parent));
      res.json({
        path: absolutePath,
        parent: parentAllowed ? parent : null,
        dirs,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const status = /restricted|traversal/i.test(message) ? 403 : 400;
      res.status(status).json({ error: message });
    }
  });

  // Remote log relay — client errors are printed to the server console
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

  // Copilot health check — reports whether any active session is connected
  apiRouter.get('/copilot-health', (_req, res) => {
    const health = checkCopilotHealth();
    res.json(health);
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

  // ─── Plugin Management Routes ──────────────────────────────────────────────

  // GET /api/plugins/installed — list sandbox-installed plugins with enriched metadata
  apiRouter.get('/plugins/installed', async (_req, res) => {
    try {
      const manager = await getPluginManager();
      await manager.refresh();
      res.json({ plugins: manager.list() });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.get('/plugins/marketplaces', async (_req, res) => {
    const result = await listPluginMarketplaces();
    if (!result.success) {
      res.status(500).json({ error: result.error ?? 'Failed to list plugin marketplaces' });
      return;
    }
    res.json({
      marketplaces: (result.data ?? []).map(marketplace => ({
        slug: marketplace.name,
        name: marketplace.name,
        kind: marketplace.kind,
        sourceKind: marketplace.sourceKind,
        source: marketplace.source,
        isBuiltIn: marketplace.kind === 'builtin',
        isOwn: marketplace.name === OCA_MARKETPLACE_NAME,
        registeredKey: marketplace.kind === 'registered' ? marketplace.name : null,
      })),
    });
  });

  apiRouter.get('/plugins/browse/:marketplace', async (req, res) => {
    const result = await browsePluginMarketplace(req.params.marketplace);
    if (!result.success) {
      res.status(500).json({ error: result.error ?? 'Failed to browse plugin marketplace' });
      return;
    }
    const manager = await getPluginManager();
    const installed = new Set(manager.list().map(plugin => `${plugin.name}@${plugin.marketplace}`));
    res.json({
      marketplace: req.params.marketplace,
      plugins: (result.data ?? []).map(plugin => ({
        ...plugin,
        installed: installed.has(`${plugin.name}@${plugin.marketplace}`),
      })),
    });
  });

  apiRouter.get('/plugins/:name/details', async (req, res) => {
    const manager = await getPluginManager();
    const plugin = manager.list().find(p => p.name === req.params.name);
    if (!plugin) {
      res.status(404).json({ error: `Plugin '${req.params.name}' not found` });
      return;
    }
    res.json(plugin);
  });

  apiRouter.post('/plugins/install', async (req, res) => {
    const spec = typeof req.body?.spec === 'string' ? req.body.spec.trim() : '';
    if (!spec) {
      res.status(400).json({ success: false, message: 'Plugin spec is required.' });
      return;
    }
    const result = await installPlugin(spec);
    await (await getPluginManager()).refresh();
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || `Installed ${spec}` : result.error,
      plugin: spec,
    });
  });

  apiRouter.post('/plugins/uninstall', async (req, res) => {
    const name = typeof req.body?.name === 'string' ? req.body.name.trim() : '';
    const marketplace = typeof req.body?.marketplace === 'string' ? req.body.marketplace.trim() : undefined;
    if (!name) {
      res.status(400).json({ success: false, message: 'Plugin name is required.' });
      return;
    }
    const result = await uninstallPlugin(name, marketplace);
    await (await getPluginManager()).refresh();
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || `Uninstalled ${name}` : result.error,
      plugin: name,
    });
  });

  apiRouter.post('/plugins/update', async (req, res) => {
    const name = typeof req.body?.name === 'string' ? req.body.name.trim() : '';
    const marketplace = typeof req.body?.marketplace === 'string' ? req.body.marketplace.trim() : undefined;
    const result = await updatePlugin(name, marketplace);
    await (await getPluginManager()).refresh();
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || (name ? `Updated ${name}` : 'Updated plugins') : result.error,
      plugin: name || undefined,
    });
  });

  apiRouter.post('/plugins/update-all', async (_req, res) => {
    const result = await updatePlugin();
    await (await getPluginManager()).refresh();
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || 'Updated plugins' : result.error,
    });
  });

  apiRouter.post('/plugins/marketplace/add', async (req, res) => {
    const spec = typeof req.body?.spec === 'string' ? req.body.spec.trim() : '';
    if (!spec) {
      res.status(400).json({ success: false, message: 'Marketplace source is required.' });
      return;
    }
    const result = await addPluginMarketplace(spec);
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || `Added marketplace ${spec}` : result.error,
    });
  });

  apiRouter.post('/plugins/marketplace/remove', async (req, res) => {
    const name =
      (typeof req.body?.registeredKey === 'string' && req.body.registeredKey.trim()) ||
      (typeof req.body?.slug === 'string' && req.body.slug.trim()) ||
      (typeof req.body?.name === 'string' && req.body.name.trim());
    if (!name) {
      res.status(400).json({ success: false, message: 'Marketplace name is required.' });
      return;
    }
    if (name === OCA_MARKETPLACE_NAME) {
      res.status(400).json({ success: false, message: 'Cannot remove the office-coding-agent marketplace.' });
      return;
    }
    const result = await removePluginMarketplace(name, { force: true });
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || `Removed marketplace ${name}` : result.error,
    });
  });

  apiRouter.post('/plugins/marketplace/update', async (req, res) => {
    const name = typeof req.body?.name === 'string' ? req.body.name.trim() : undefined;
    const result = await updatePluginMarketplace(name);
    res.status(result.success ? 200 : 500).json({
      success: result.success,
      message: result.success ? result.stdout.trim() || `Updated marketplace${name ? ` ${name}` : 's'}` : result.error,
    });
  });

  // GET /api/mcp-servers — all available MCP server configs (bundled + from installed plugins)
  apiRouter.get('/mcp-servers', async (_req, res) => {
    try {
      const manager = await getPluginManager();
      const pluginServers = manager.getMcpServerConfigs();
      const bundledEndpoints = new Set(
        BUNDLED_MCP_SERVERS.map(s =>
          s.transport === 'stdio' ? `stdio:${s.command}` : `url:${(s.url ?? '').replace(/\/$/, '')}`
        )
      );
      const bundledNames = new Set(BUNDLED_MCP_SERVERS.map(s => s.name));
      const endpointKey = s =>
        s.transport === 'stdio' ? `stdio:${s.command}` : `url:${(s.url ?? '').replace(/\/$/, '')}`;
      const dedupedPluginServers = pluginServers.filter(
        s => !bundledNames.has(s.name) && !bundledEndpoints.has(endpointKey(s))
      );
      res.json({ servers: [...BUNDLED_MCP_SERVERS, ...dedupedPluginServers] });
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

    // Auto-setup plugins in background (non-blocking)
    setTimeout(autoSetupPlugins, 500);
  });
}

createServer().catch(err => {
  console.error('Server startup error:', err);
  process.exit(1);
});
