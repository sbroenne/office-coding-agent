import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import { getCopilotHomeDir } from './pluginPaths.mjs';

const DEFAULT_TIMEOUT_MS = 120_000;
const MAX_OUTPUT_BYTES = 64 * 1024;

function getCliCommand() {
  return process.env.COPILOT_CLI_PATH || (process.platform === 'win32' ? 'copilot.cmd' : 'copilot');
}

export function installPlugin(spec, timeoutMs = DEFAULT_TIMEOUT_MS) {
  return runPluginCli(['install', spec], timeoutMs);
}

export function uninstallPlugin(name, marketplace, timeoutMs = DEFAULT_TIMEOUT_MS) {
  return runPluginCli(['uninstall', marketplace ? `${name}@${marketplace}` : name], timeoutMs);
}

export function updatePlugin(name, marketplace, timeoutMs = DEFAULT_TIMEOUT_MS) {
  const args = name ? ['update', marketplace ? `${name}@${marketplace}` : name] : ['update'];
  return runPluginCli(args, timeoutMs);
}

export function listPluginMarketplaces(timeoutMs = DEFAULT_TIMEOUT_MS) {
  return runPluginCliAndParse(
    ['marketplace', 'list'],
    timeoutMs,
    parseMarketplaceListOutput,
    'marketplace list'
  );
}

export function browsePluginMarketplace(name, timeoutMs = DEFAULT_TIMEOUT_MS) {
  return runPluginCliAndParse(
    ['marketplace', 'browse', name],
    timeoutMs,
    stdout => parseMarketplaceBrowseOutput(stdout, name),
    `marketplace browse ${name}`
  );
}

export function addPluginMarketplace(source, timeoutMs = DEFAULT_TIMEOUT_MS) {
  return runPluginCli(['marketplace', 'add', source], timeoutMs);
}

export function removePluginMarketplace(name, options = {}, timeoutMs = DEFAULT_TIMEOUT_MS) {
  const args = ['marketplace', 'remove', name];
  if (options.force) args.push('--force');
  return runPluginCli(args, timeoutMs);
}

export function updatePluginMarketplace(name, timeoutMs = DEFAULT_TIMEOUT_MS) {
  const args = ['marketplace', 'update'];
  if (name) args.push(name);
  return runPluginCli(args, timeoutMs);
}

export function parseMarketplaceListOutput(stdout) {
  const normalized = normalizeMarketplaceOutput(stdout);
  const marketplaces = [];
  let currentKind = null;

  for (const rawLine of normalized.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    if (line.includes('Included with GitHub Copilot')) {
      currentKind = 'builtin';
      continue;
    }
    if (line.startsWith('Registered marketplaces')) {
      currentKind = 'registered';
      continue;
    }
    const match = line.match(/([A-Za-z0-9][\w.-]*)\s+\(([^:]+):\s*(.+)\)$/);
    if (!match || !currentKind) continue;
    const [, name, sourceKind, source] = match;
    marketplaces.push({
      name,
      kind: currentKind,
      sourceKind: sourceKind.trim(),
      source: source.trim(),
    });
  }

  return marketplaces;
}

export function parseMarketplaceBrowseOutput(stdout, fallbackMarketplace) {
  const normalized = normalizeMarketplaceOutput(stdout);
  const entries = [];
  let marketplace = fallbackMarketplace;

  for (const rawLine of normalized.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    const headingMatch = line.match(/^Plugins in "(.+)":$/);
    if (headingMatch) {
      marketplace = headingMatch[1];
      continue;
    }
    if (line.startsWith('Install with:')) continue;
    const match = line.match(/([A-Za-z0-9][\w.-]*)\s+-\s+(.+)$/);
    if (!match) continue;
    const [, name, description] = match;
    entries.push({ name, marketplace, description: description.trim() });
  }

  return entries;
}

function normalizeMarketplaceOutput(stdout) {
  return stdout.replace(/\)(?=Registered marketplaces:)/g, ')\n').replace(/(?<=\S)(?=Install with:)/g, '\n');
}

async function runPluginCliAndParse(args, timeoutMs, parse, description) {
  const result = await runPluginCli(args, timeoutMs);
  if (!result.success) return result;
  try {
    return { ...result, data: parse(result.stdout) };
  } catch (err) {
    return {
      ...result,
      success: false,
      error: err instanceof Error ? `${description}: ${err.message}` : `${description} failed`,
    };
  }
}

export function runPluginCli(args, timeoutMs = DEFAULT_TIMEOUT_MS) {
  return new Promise(resolve => {
    const copilotHome = getCopilotHomeDir();
    try {
      mkdirSync(copilotHome, { recursive: true });
    } catch (err) {
      resolve({
        success: false,
        stdout: '',
        stderr: '',
        error: err instanceof Error ? err.message : String(err),
      });
      return;
    }

    const proc = spawn(getCliCommand(), ['plugin', ...args], {
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
      env: { ...process.env, COPILOT_HOME: copilotHome },
      shell: process.platform === 'win32',
    });

    let stdout = '';
    let stderr = '';
    let timedOut = false;

    const append = (current, chunk) => {
      if (current.length >= MAX_OUTPUT_BYTES) return current;
      return (current + chunk.toString('utf-8')).slice(0, MAX_OUTPUT_BYTES);
    };

    proc.stdout.on('data', chunk => {
      stdout = append(stdout, chunk);
    });
    proc.stderr.on('data', chunk => {
      stderr = append(stderr, chunk);
    });

    const timer = setTimeout(() => {
      timedOut = true;
      proc.kill();
    }, timeoutMs);

    proc.on('error', err => {
      clearTimeout(timer);
      resolve({ success: false, stdout, stderr, error: err.message });
    });

    proc.on('close', code => {
      clearTimeout(timer);
      if (code === 0 && !timedOut) {
        resolve({ success: true, stdout, stderr });
        return;
      }
      const tail = (stderr || stdout).trim().slice(-500);
      resolve({
        success: false,
        stdout,
        stderr,
        error: timedOut
          ? `copilot plugin ${args.join(' ')} timed out after ${timeoutMs}ms`
          : tail || `copilot plugin ${args.join(' ')} exited with code ${code ?? 'null'}`,
      });
    });
  });
}
