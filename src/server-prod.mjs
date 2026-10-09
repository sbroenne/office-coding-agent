import express from 'express';
import cors from 'cors';
import https from 'node:https';
import path from 'node:path';
import fs from 'node:fs';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { setupCopilotProxy, checkCopilotHealth } from './copilotProxy.mjs';
import { ensureOfficeCliPlugins } from './plugins/cliPluginBootstrap.mjs';
import { getCliSlashItems } from './plugins/cliSlashItems.mjs';
import { getCliMcpServers, getMcpServerSummaries } from './plugins/cliMcpServers.mjs';
import {
  getBrowseRoots,
  isAllowedOrigin,
  isPathWithinRoot,
  isTrustedRequestOrigin,
  resolveBrowsePath,
} from './serverSecurity.mjs';

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
    tester.listen(port, 'localhost');
  });
}

export async function createServer() {
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

  const requireTrustedLocalAccess = (req, res, next) => {
    const remoteAddress = req.socket?.remoteAddress;
    if (isTrustedRequestOrigin(req.headers.origin, remoteAddress)) {
      next();
      return;
    }

    res.status(403).json({ error: 'This endpoint is only available to the local add-in.' });
  };

  const apiRouter = express.Router();
  apiRouter.use(express.json({ limit: '50mb' }));

  apiRouter.get('/hello', (_req, res) => {
    res.json({ message: 'Copilot proxy running', timestamp: new Date().toISOString() });
  });

  apiRouter.get('/ping', (_req, res) => {
    res.json({ ok: true });
  });

  apiRouter.get('/env', requireTrustedLocalAccess, (_req, res) => {
    res.json({
      platform: process.platform,
      nodeEnv: process.env.NODE_ENV ?? 'production',
      browseRestricted: true,
    });
  });

  apiRouter.get('/slash-items', requireTrustedLocalAccess, async (_req, res) => {
    try {
      res.json(await getCliSlashItems());
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
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
      const parentAllowed =
        parent !== absolutePath && browseRoots.some(root => isPathWithinRoot(root, parent));
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

  // GET /api/mcp-servers — MCP server configs from the user's Copilot CLI config.
  apiRouter.get('/mcp-servers', requireTrustedLocalAccess, async (_req, res) => {
    const result = await getCliMcpServers();
    if (result.error) {
      console.warn(`[mcp] Failed to load Copilot CLI MCP servers: ${result.error}`);
    }
    res.json({
      servers: getMcpServerSummaries(result.servers),
      ...(result.error ? { error: 'Some MCP server settings could not be loaded.' } : {}),
    });
  });

  app.use('/api', apiRouter);
  app.get('/ping', (_req, res) => res.json({ ok: true }));

  const devCerts = await import('office-addin-dev-certs');
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const httpsServer = https.createServer(httpsOptions, app);

  await ensureOfficeCliPlugins();
  setupCopilotProxy(httpsServer);
  const mcpStartup = await getCliMcpServers();
  if (mcpStartup.error) {
    console.warn(`[mcp] Failed to load Copilot CLI MCP servers: ${mcpStartup.error}`);
  }

  const distDir = path.resolve(__dirname, '..', 'dist');
  app.use(express.static(distDir));
  app.get('*path', (_req, res) => {
    res.sendFile(path.join(distDir, 'taskpane.html'));
  });

  await new Promise(resolve => {
    httpsServer.listen(PORT, 'localhost', () => {
      console.log(
        `\n  Copilot Office Add-in production server running on https://localhost:${PORT}`
      );
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
