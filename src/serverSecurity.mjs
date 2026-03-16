import path from 'node:path';
import fs from 'node:fs/promises';
import os from 'node:os';

const LOOPBACK_HOST_RE = /^(localhost|127\.0\.0\.1|::1)$/i;
const TRUSTED_HTTPS_HOST_RE = /^(?:[a-z0-9-]+\.)*(officeapps\.live\.com|office\.com|microsoft\.com)$/i;
const PATH_TRAVERSAL_RE = /(^|[\\/])\.\.([\\/]|$)/;

export function isAllowedOrigin(origin) {
  if (!origin || typeof origin !== 'string') return false;

  try {
    const parsed = new URL(origin);
    const protocol = parsed.protocol.toLowerCase();
    const host = parsed.hostname.toLowerCase();

    if (LOOPBACK_HOST_RE.test(host)) {
      return protocol === 'https:' || protocol === 'http:';
    }

    return protocol === 'https:' && TRUSTED_HTTPS_HOST_RE.test(host);
  } catch {
    return false;
  }
}

export function isLoopbackAddress(address) {
  if (!address || typeof address !== 'string') return false;

  const normalized = address.toLowerCase();
  return (
    normalized === '::1' ||
    normalized === '127.0.0.1' ||
    normalized === '::ffff:127.0.0.1' ||
    normalized === '::ffff:7f00:1'
  );
}

export function isTrustedRequestOrigin(origin, remoteAddress) {
  if (origin) {
    return isAllowedOrigin(origin);
  }

  return isLoopbackAddress(remoteAddress);
}

export async function getBrowseRoots() {
  const candidates = [process.cwd(), os.homedir()];
  const resolvedRoots = await Promise.all(
    candidates.map(async candidate => {
      try {
        return await fs.realpath(candidate);
      } catch {
        return path.resolve(candidate);
      }
    })
  );

  return [...new Set(resolvedRoots)];
}

export function isPathWithinRoot(rootPath, targetPath) {
  const relativePath = path.relative(rootPath, targetPath);
  return relativePath === '' || (!relativePath.startsWith('..') && !path.isAbsolute(relativePath));
}

export async function resolveBrowsePath(requestedPath, allowedRoots) {
  if (typeof requestedPath === 'string' && PATH_TRAVERSAL_RE.test(requestedPath)) {
    throw new Error('Path traversal is not allowed.');
  }

  const candidatePath = typeof requestedPath === 'string' && requestedPath.trim()
    ? requestedPath
    : process.cwd();
  const realPath = await fs.realpath(path.resolve(candidatePath));

  const isAllowed = allowedRoots.some(root => isPathWithinRoot(root, realPath));
  if (!isAllowed) {
    throw new Error('Browsing is restricted to approved directories.');
  }

  return realPath;
}
