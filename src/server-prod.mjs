import express from 'express';
import cors from 'cors';
import https from 'node:https';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { execSync } from 'node:child_process';
import { setupCopilotProxy, checkCopilotHealth } from './copilotProxy.mjs';

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
    } catch { /* ignore */ }
  }

  const skillsDir = path.join(pluginDir, 'skills');
  if (fs.existsSync(skillsDir)) {
    try {
      const entries = fs.readdirSync(skillsDir, { withFileTypes: true, recursive: true });
      for (const entry of entries) {
        if (entry.isFile() && entry.name === 'SKILL.md') {
          const parentPath = entry.parentPath || entry.path || '';
          result.skillCount++;
          result.skillNames.push(path.basename(parentPath));
        }
      }
    } catch { /* ignore */ }
  }

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

// ─── Auto-Setup: register marketplace and install host plugins ──────────────

const OCA_MARKETPLACE = 'sbroenne/office-coding-agent-plugins';
const OCA_MARKETPLACE_NAME = 'office-coding-agent';
const HOST_PLUGINS = ['office-excel', 'office-powerpoint', 'office-word', 'office-outlook'];

function autoSetupPlugins() {
  try {
    const config = readCopilotConfig();
    const marketplaces = config.marketplaces || {};
    const installed = config.installed_plugins || [];

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

  // ─── Plugin Management Routes ──────────────────────────────────────────────

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

  apiRouter.get('/plugins/browse/:marketplace', (_req, res) => {
    try {
      const { marketplace } = _req.params;
      const cacheDir = path.join(os.homedir(), '.copilot', 'marketplace-cache');

      if (!fs.existsSync(cacheDir)) {
        res.status(404).json({ error: 'Marketplace cache not found' });
        return;
      }

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

      const config = readCopilotConfig();
      const installedNames = new Set((config.installed_plugins || []).map(p => p.name));
      const plugins = manifest.plugins.map(plugin => ({
        ...plugin,
        installed: installedNames.has(plugin.name),
      }));

      res.json({ marketplace: manifest.name || marketDir.name, plugins });
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

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

  apiRouter.post('/plugins/install', (req, res) => {
    try {
      const { spec } = req.body;
      if (!spec || typeof spec !== 'string') { res.status(400).json({ error: 'Missing required field: spec' }); return; }
      const result = runCopilotCommand(`install ${spec}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/plugins/uninstall', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') { res.status(400).json({ error: 'Missing required field: name' }); return; }
      const result = runCopilotCommand(`uninstall ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/plugins/enable', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') { res.status(400).json({ error: 'Missing required field: name' }); return; }
      const result = runCopilotCommand(`enable ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/plugins/disable', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') { res.status(400).json({ error: 'Missing required field: name' }); return; }
      const result = runCopilotCommand(`disable ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/plugins/update', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') { res.status(400).json({ error: 'Missing required field: name' }); return; }
      const result = runCopilotCommand(`update ${name}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/plugins/marketplace/add', (req, res) => {
    try {
      const { spec } = req.body;
      if (!spec || typeof spec !== 'string') { res.status(400).json({ error: 'Missing required field: spec' }); return; }
      const result = runCopilotCommand(`marketplace add ${spec}`);
      res.status(result.success ? 200 : 500).json(result);
    } catch (error) {
      res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
    }
  });

  apiRouter.post('/plugins/marketplace/remove', (req, res) => {
    try {
      const { name } = req.body;
      if (!name || typeof name !== 'string') { res.status(400).json({ error: 'Missing required field: name' }); return; }
      const result = runCopilotCommand(`marketplace remove ${name}`);
      res.status(result.success ? 200 : 500).json(result);
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
            const server = { name, description: `From plugin: ${plugin.name}`, ...cfg };
            if (cfg.command) server.transport = 'stdio';
            else if (cfg.url) server.transport = 'http';
            pluginServers.push(server);
          }
        } catch { /* skip malformed .mcp.json */ }
      }

      // Bundled servers take priority — drop any plugin server with the same name
      const bundledNames = new Set(BUNDLED.map(s => s.name));
      const dedupedPluginServers = pluginServers.filter(s => !bundledNames.has(s.name));
      res.json({ servers: [...BUNDLED, ...dedupedPluginServers] });
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
      console.log(
        `\n  Copilot Office Add-in production server running on https://localhost:${PORT}`
      );
      console.log(`  API: https://localhost:${PORT}/api\n`);
      resolve(undefined);
    });
  });

  // Auto-setup plugins in background (non-blocking)
  setTimeout(autoSetupPlugins, 500);

  return httpsServer;
}

const isMainModule = process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1]);

if (isMainModule) {
  createServer().catch(err => {
    console.error('Server startup error:', err);
    process.exit(1);
  });
}
