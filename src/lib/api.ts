/**
 * Returns the base URL for the local proxy API server.
 *
 * The API server always runs on localhost:3000 — even when the taskpane HTML
 * is served from a remote origin (e.g. GitHub Pages). All fetch('/api/...')
 * calls must use this base so they reach the local server, not the CDN.
 */
export function getLocalApiBase(): string {
  if (typeof window === 'undefined') return 'https://localhost:3000';
  const { hostname, protocol, host } = window.location;
  // Non-localhost origin (GitHub Pages, staging) → always target localhost:3000
  if (hostname !== 'localhost' && hostname !== '127.0.0.1') {
    return 'https://localhost:3000';
  }
  return `${protocol}//${host}`;
}
