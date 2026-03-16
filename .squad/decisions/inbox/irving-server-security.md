# Irving — Server security hardening decisions

## Context
Security review flagged overly-permissive localhost API exposure in `src/server.mjs` and unchecked WebSocket upgrades in `src/copilotProxy.mjs`.

## Decisions made

1. **Restrict API origins to trusted callers**
   - Allow localhost variants (`https://localhost:*`, `http://localhost:*`, `127.0.0.1`, `::1`) for local development and same-machine tooling.
   - Allow trusted Microsoft web origins (`https://*.officeapps.live.com`, `https://*.office.com`, `https://*.microsoft.com`) for Office-hosted task pane scenarios.
   - Reject unexpected `Origin` headers on `/api/*` with HTTP 403 instead of silently serving the request.

2. **Gate sensitive localhost APIs by trusted origin or loopback caller**
   - `/api/env` and `/api/browse` now require either a trusted `Origin` header or a no-origin loopback request.
   - This treats the local add-in/browser as the only supported caller while preserving same-machine diagnostics and automated tests.

3. **Reduce `/api/env` data to non-sensitive metadata**
   - Removed filesystem path disclosure (`cwd`, `home`).
   - Kept only non-sensitive runtime hints (`platform`, `nodeEnv`, `browseRestricted`).
   - Updated the Permission Manager UI to bootstrap directly from `/api/browse`, so it no longer needs path data from `/api/env`.

4. **Constrain `/api/browse` to canonical approved roots**
   - Approved roots are the project working directory and the user home directory.
   - Requests are resolved through `realpath()` before authorization so prefix tricks and symlink escapes do not bypass the allowlist.
   - Any request containing `..` path traversal segments is rejected.

5. **Validate WebSocket upgrade origins**
   - `/api/copilot` now rejects untrusted upgrade requests with HTTP 403 before `ws.handleUpgrade()` runs.
   - No-origin loopback upgrades remain allowed for local tooling, but browser-origin upgrades must match the trusted allowlist.

## Validation
- `npm run build:dev`
- `npm run test:integration`
- Manual negative checks:
  - bad REST origin => 403
  - `/api/browse?path=..\\` => 403
  - bad WebSocket origin => 403

## Follow-up notes
- The allowlist is intentionally practical, not exhaustive. If future Office hosts use different origins, extend `src/serverSecurity.mjs` rather than reopening wildcard CORS.
- This change hardens both `src/server.mjs` and `src/server-prod.mjs` to avoid dev/prod security drift.
