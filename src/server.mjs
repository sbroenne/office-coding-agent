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

  // GET /api/plugins/marketplaces — list available plugin marketplaces
  apiRouter.get('/plugins/marketplaces', (_req, res) => {
    try {
      const builtInSlugs = ['copilot-plugins', 'awesome-copilot'];
      const cacheDir = path.join(os.homedir(), '.copilot', 'marketplace-cache');
      const marketplaces = [];

      if (fs.existsSync(cacheDir)) {
        const entries = fs.readdirSync(cacheDir, { withFileTypes: true });
        for (const entry of entries) {
          if (!entry.isDirectory()) continue;
          const dirPath = path.join(cacheDir, entry.name);
          const manifest = findMarketplaceManifest(dirPath);
          const isBuiltIn = builtInSlugs.some(slug => entry.name.includes(slug));
          marketplaces.push({
            slug: entry.name,
            name: manifest?.name || entry.name,
            owner: manifest?.owner || null,
            description: manifest?.description || null,
            pluginCount: Array.isArray(manifest?.plugins) ? manifest.plugins.length : 0,
            isBuiltIn,
          });
        }
      }

      res.json({ marketplaces });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // GET /api/plugins/browse/:marketplace — browse plugins in a marketplace
  apiRouter.get('/plugins/browse/:marketplace', (_req, res) => {
    try {
      const { marketplace } = _req.params;
      const cacheDir = path.join(os.homedir(), '.copilot', 'marketplace-cache');

      if (!fs.existsSync(cacheDir)) {
        res.status(404).json({ error: 'Marketplace cache not found' });
        return;
      }

      // Find the marketplace directory (exact match or contains the slug)
      const entries = fs.readdirSync(cacheDir, { withFileTypes: true });
      const marketDir = entries.find(
        e => e.isDirectory() && (e.name === marketplace || e.name.includes(marketplace))
      );
      if (!marketDir) {
        res.status(404).json({ error: `Marketplace "${marketplace}" not found` });
        return;
      }

      const manifest = findMarketplaceManifest(path.join(cacheDir, marketDir.name));
      if (!manifest || !Array.isArray(manifest.plugins)) {
        res.status(404).json({ error: 'Marketplace manifest not found or has no plugins' });
        return;
      }

      // Cross-reference with installed plugins
      const config = readCopilotConfig();
      const installedNames = new Set(
        (config.installed_plugins || []).map(p => p.name)
      );

      const plugins = manifest.plugins.map(plugin => ({
        ...plugin,
        installed: installedNames.has(plugin.name),
      }));

      res.json({
        marketplace: manifest.name || marketDir.name,
        plugins,
      });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // GET /api/plugins/:name/details — full detail for a single installed plugin
  apiRouter.get('/plugins/:name/details', (_req, res) => {
    try {
      const { name } = _req.params;
      const config = readCopilotConfig();
      const plugin = (config.installed_plugins || []).find(p => p.name === name);

      if (!plugin) {
        res.status(404).json({ error: `Plugin "${name}" is not installed` });
        return;
      }

      let manifest = null;
      let components = null;
      if (plugin.cache_path && fs.existsSync(plugin.cache_path)) {
        manifest = findPluginManifest(plugin.cache_path);
        components = countPluginComponents(plugin.cache_path, manifest);
      }

      res.json({ ...plugin, manifest, components });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/install — install a plugin by spec
  apiRouter.post('/plugins/install', (req, res) => {
    try {
      const { spec } = req.body;
      if (!spec || typeof spec !== 'string') {
        res.status(400).json({ error: 'Missing required field: spec' });
        return;
      }
      const result = runCopilotCommand(`install ${spec}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/uninstall — uninstall a plugin by name
  apiRouter.post('/plugins/uninstall', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') {
        res.status(400).json({ error: 'Missing required field: name' });
        return;
      }
      const result = runCopilotCommand(`uninstall ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/enable — enable a plugin by name
  apiRouter.post('/plugins/enable', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') {
        res.status(400).json({ error: 'Missing required field: name' });
        return;
      }
      const result = runCopilotCommand(`enable ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/disable — disable a plugin by name
  apiRouter.post('/plugins/disable', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') {
        res.status(400).json({ error: 'Missing required field: name' });
        return;
      }
      const result = runCopilotCommand(`disable ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/update — update a plugin by name
  apiRouter.post('/plugins/update', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') {
        res.status(400).json({ error: 'Missing required field: name' });
        return;
      }
      const result = runCopilotCommand(`update ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/marketplace/add — add a marketplace by spec
  apiRouter.post('/plugins/marketplace/add', (req, res) => {
    try {
      const { spec } = req.body;
      if (!spec || typeof spec !== 'string') {
        res.status(400).json({ error: 'Missing required field: spec' });
        return;
      }
      const result = runCopilotCommand(`marketplace add ${spec}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  // POST /api/plugins/marketplace/remove — remove a marketplace by name
  apiRouter.post('/plugins/marketplace/remove', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') {
        res.status(400).json({ error: 'Missing required field: name' });
        return;
      }
      const result = runCopilotCommand(`marketplace remove ${name}`);
      res.status(result.success ? 200 : 500).json(result);
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
