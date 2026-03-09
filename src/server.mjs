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
import { execSync } from 'node:child_process';
import { setupCopilotProxy, checkCopilotHealth } from './copilotProxy.mjs';
import { listMarketplaces, removeMarketplace as removeMarketplaceSvc } from './marketplaceService.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = 3000;
const isDev = process.env.NODE_ENV !== 'production';

// ─── Copilot Plugin Helpers ────────────────────────────────────────────────

function getCopilotConfigPath() {
  return path.join(os.homedir(), '.copilot', 'config.json');
}

function readCopilotConfig() {
  const configPath = getCopilotConfigPath();
  if (!fs.existsSync(configPath)) return { installed_plugins: [] };
  try {
    return JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  } catch {
    return { installed_plugins: [] };
  }
}

function findPluginManifest(pluginDir) {
  const candidates = [
    path.join(pluginDir, 'plugin.json'),
    path.join(pluginDir, '.github', 'plugin', 'plugin.json'),
    path.join(pluginDir, '.claude-plugin', 'plugin.json'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      try {
        return JSON.parse(fs.readFileSync(p, 'utf-8'));
      } catch {
        continue;
      }
    }
  }
  return null;
}

function findMarketplaceManifest(cacheDir) {
  const candidates = [
    path.join(cacheDir, '.github', 'plugin', 'marketplace.json'),
    path.join(cacheDir, '.claude-plugin', 'marketplace.json'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      try {
        return JSON.parse(fs.readFileSync(p, 'utf-8'));
      } catch {
        continue;
      }
    }
  }
  return null;
}

function countPluginComponents(pluginDir, manifest) {
  const result = {
    agentCount: 0, agentNames: [],
    skillCount: 0, skillNames: [],
    mcpServerCount: 0, mcpServerNames: [],
    hookCount: 0, commandCount: 0,
  };

  // Count agents from agents/ directory (*.agent.md files)
  const agentsDir = path.join(pluginDir, 'agents');
  if (fs.existsSync(agentsDir)) {
    try {
      const entries = fs.readdirSync(agentsDir, { withFileTypes: true, recursive: true });
      for (const entry of entries) {
        if (entry.isFile() && entry.name.endsWith('.agent.md')) {
          result.agentCount++;
          result.agentNames.push(entry.name.replace(/\.agent\.md$/, ''));
        }
      }
    } catch { /* ignore unreadable dirs */ }
  }

  // Count skills from skills/ directory (SKILL.md files)
  const skillsDir = path.join(pluginDir, 'skills');
  if (fs.existsSync(skillsDir)) {
    try {
      const entries = fs.readdirSync(skillsDir, { withFileTypes: true, recursive: true });
      for (const entry of entries) {
        if (entry.isFile() && entry.name === 'SKILL.md') {
          // Use parent directory name as skill name
          const parentPath = entry.parentPath || entry.path || '';
          result.skillCount++;
          result.skillNames.push(path.basename(parentPath));
        }
      }
    } catch { /* ignore */ }
  }

  // Count MCP servers from .mcp.json
  const mcpJsonPath = path.join(pluginDir, '.mcp.json');
  if (fs.existsSync(mcpJsonPath)) {
    try {
      const mcpConfig = JSON.parse(fs.readFileSync(mcpJsonPath, 'utf-8'));
      const servers = mcpConfig.mcpServers || mcpConfig.servers || {};
      const names = Object.keys(servers);
      result.mcpServerCount = names.length;
      result.mcpServerNames = names;
    } catch { /* ignore */ }
  }

  // Count hooks and commands from manifest
  if (manifest) {
    if (Array.isArray(manifest.hooks)) result.hookCount = manifest.hooks.length;
    if (Array.isArray(manifest.commands)) result.commandCount = manifest.commands.length;
  }

  return result;
}

function runCopilotCommand(args) {
  try {
    const output = execSync(`copilot plugin ${args}`, {
      encoding: 'utf-8',
      timeout: 60000,
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    return { success: true, message: output.trim() };
  } catch (err) {
    return { success: false, message: err.stderr?.trim() || err.message };
  }
}

// ─── Auto-Setup: register our marketplace and install host plugins ──────────

const OCA_MARKETPLACE = 'sbroenne/office-coding-agent-plugins';
const OCA_MARKETPLACE_NAME = 'office-coding-agent';
const HOST_PLUGINS = ['office-excel', 'office-powerpoint', 'office-word', 'office-outlook'];

function autoSetupPlugins() {
  try {
    const config = readCopilotConfig();
    const marketplaces = config.marketplaces || {};
    const installed = config.installed_plugins || [];

    // 1. Register our marketplace if not already registered
    const hasMarketplace = OCA_MARKETPLACE_NAME in marketplaces;
    if (!hasMarketplace) {
      console.log(`  [auto-setup] Registering marketplace: ${OCA_MARKETPLACE}`);
      const result = runCopilotCommand(`marketplace add ${OCA_MARKETPLACE}`);
      if (result.success) {
        console.log(`  [auto-setup] ✓ Marketplace registered`);
      } else {
        console.warn(`  [auto-setup] ⚠ Marketplace registration failed: ${result.message}`);
      }
    }

    // 2. Install host plugins if missing
    for (const pluginName of HOST_PLUGINS) {
      const isInstalled = Array.isArray(installed)
        ? installed.some(p => p.name === pluginName)
        : false;
      if (!isInstalled) {
        console.log(`  [auto-setup] Installing plugin: ${pluginName}`);
        const result = runCopilotCommand(`install ${pluginName}@${OCA_MARKETPLACE_NAME}`);
        if (result.success) {
          console.log(`  [auto-setup] ✓ ${pluginName} installed`);
        } else {
          console.warn(`  [auto-setup] ⚠ ${pluginName} install failed: ${result.message}`);
        }
      }
    }

    console.log('  [auto-setup] Plugin setup complete');
  } catch (err) {
    console.warn(`  [auto-setup] Auto-setup failed (non-fatal): ${err.message}`);
  }
}

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

      // Security: restrict browsing to home directory or current working directory
      const homeDir = os.homedir();
      const cwdDir = process.cwd();
      const isAllowed = absolutePath.startsWith(homeDir) || absolutePath.startsWith(cwdDir);
      if (!isAllowed) {
        res.status(403).json({ error: 'Browsing is restricted to the home or project directory.' });
        return;
      }

      const entries = await fs.promises.readdir(absolutePath, { withFileTypes: true });
      const dirs = entries
        .filter(entry => entry.isDirectory())
        .map(entry => entry.name)
        .sort((a, b) => a.localeCompare(b));
      const parent = path.dirname(absolutePath);
      const parentAllowed = parent !== absolutePath && (parent.startsWith(homeDir) || parent.startsWith(cwdDir));
      res.json({
        path: absolutePath,
        parent: parentAllowed ? parent : null,
        dirs,
      });
    } catch (error) {
      res.status(400).json({ error: error instanceof Error ? error.message : String(error) });
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

  // GET /api/plugins/installed — list installed plugins with enriched metadata
  apiRouter.get('/plugins/installed', (_req, res) => {
    try {
      const config = readCopilotConfig();
      const plugins = (config.installed_plugins || []).map(plugin => {
        let manifest = null;
        let components = null;
        if (plugin.cache_path && fs.existsSync(plugin.cache_path)) {
          manifest = findPluginManifest(plugin.cache_path);
          components = countPluginComponents(plugin.cache_path, manifest);
        }
        return { ...plugin, manifest, components };
      });
      res.json({ plugins });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // GET /api/mcp-servers — all available MCP server configs (bundled + from installed plugins)
  apiRouter.get('/mcp-servers', (_req, res) => {
    try {
      const BUNDLED = [
        { name: 'workiq', description: 'Microsoft 365 Copilot — emails, meetings, documents, Teams', transport: 'stdio', command: 'npx', args: ['-y', '@microsoft/workiq', 'mcp'] },
        { name: 'powerbi', description: 'Power BI — query semantic models, generate DAX, explore data', transport: 'http', url: 'https://api.fabric.microsoft.com/v1/mcp/powerbi' },
      ];

      const pluginServers = [];
      const config = readCopilotConfig();
      const installedPlugins = config.installed_plugins || [];

      for (const plugin of installedPlugins) {
        if (!plugin.enabled) continue;
        if (!plugin.cache_path || !fs.existsSync(plugin.cache_path)) continue;
        const mcpJsonPath = path.join(plugin.cache_path, '.mcp.json');
        if (!fs.existsSync(mcpJsonPath)) continue;
        try {
          const mcpConfig = JSON.parse(fs.readFileSync(mcpJsonPath, 'utf-8'));
          const servers = mcpConfig.mcpServers || mcpConfig.servers || {};
          for (const [name, cfg] of Object.entries(servers)) {
            // Normalise to McpServerConfig shape
            const server = { name, description: `From plugin: ${plugin.name}`, ...cfg };
            if (cfg.command) server.transport = 'stdio';
            else if (cfg.url) server.transport = 'http';
            pluginServers.push(server);
          }
        } catch { /* skip malformed .mcp.json */ }
      }

      // Bundled servers take priority — drop any plugin server pointing to the same
      // endpoint (same command for stdio, same url for http/sse), regardless of name.
      const bundledEndpoints = new Set(
        BUNDLED.map(s => (s.transport === 'stdio' ? `stdio:${s.command}` : `url:${(s.url ?? '').replace(/\/$/, '')}`))
      );
      const endpointKey = s =>
        s.transport === 'stdio' ? `stdio:${s.command}` : `url:${(s.url ?? '').replace(/\/$/, '')}`;
      const dedupedPluginServers = pluginServers.filter(s => !bundledEndpoints.has(endpointKey(s)));
      res.json({ servers: [...BUNDLED, ...dedupedPluginServers] });
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
