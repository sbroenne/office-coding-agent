# Sideloading Guide

> **First time?** See [GETTING_STARTED.md](../GETTING_STARTED.md) for the full setup walkthrough including authentication, proxy server startup, and add-in registration.
>
> For normal local development, use `npm run start:dev:desktop`. It starts the proxy server, waits for `https://localhost:3000`, then sideloads Excel.

This project supports three sideloading lanes:

1. **Local desktop dev (fastest)** via `localhost`
2. **Local shared folder catalog** (Windows testing flow)
3. **Staging manifest** that points to **GitHub Pages**

Main environment manifests are stored in `manifests/`:

- `manifests/manifest.dev.xml`
- `manifests/manifest.staging.xml`
- `manifests/manifest.prod.xml`

## Important Note

A shared folder catalog distributes the **manifest only**.

The add-in web app (task pane HTML/JS/CSS) must be hosted at the HTTPS URLs in the manifest (`SourceLocation`, icon URLs).

- For one-machine local dev, `https://localhost:3000` is fine.
- For testing from other machines, use `manifests/manifest.staging.xml` that points to GitHub Pages.

## Lane 1: Local Desktop Dev

```bash
npm run start:dev:desktop
npm run start:dev:desktop:ppt
npm run start:dev:desktop:word
```

If the dev server is already running, you can sideload only:

```bash
npm run start:desktop:excel
npm run start:desktop:ppt
npm run start:desktop:word
```

Or, to run through tray mode first:

```bash
npm run start:tray:excel
npm run start:tray:ppt
npm run start:tray:word
```

Note: tray startup includes a reliability fallback — if the tray server does not
become healthy on `https://localhost:3000` in time, the launcher falls back to
starting the dev server and then proceeds with host sideload.

When done:

```bash
npm run stop
```

## Lane 2: Local Shared Folder Catalog (Windows)

### Elevation requirements

- `sideload:share:setup` requires **Administrator** only when creating the SMB share.
- `sideload:share:cleanup` requires **Administrator** only when removing an existing SMB share.
- `sideload:share:trust` and `sideload:share:publish` run as normal user.

The scripts now detect missing elevation and return a clear instruction to rerun in elevated PowerShell.

### 1) Create local share

```bash
npm run sideload:share:setup
```

Default local folder: `%USERPROFILE%\OfficeAddinCatalog`  
Default share name: `OfficeAddinCatalog`

### 2) Trust catalog in Office

```bash
npm run sideload:share:trust
```

Restart Excel after trust registration.

### 3) Publish staging manifest into share

```bash
npm run sideload:share:publish
```

In Excel: **Home > Add-ins > More Add-ins > Shared Folder**, then add `manifest.staging.xml`.

### 4) Cleanup

```bash
npm run sideload:share:cleanup
```

## Lane 3: Staging on GitHub Pages

GitHub Pages deploys from `main` via `.github/workflows/pages.yml`.

Staging manifest target base URL:

- `https://sbroenne.github.io/office-coding-agent`

Committed file:

- `manifests/manifest.staging.xml`

## CLI Plugin and MCP Checklist

Use this after the add-in is loaded (desktop or staging) to verify the CLI-owned plugin and MCP flows.

1. Confirm required Office plugins are installed by the startup bootstrap:

```bash
copilot plugin list
```

2. Manage user plugins with the Copilot CLI:

```bash
copilot plugin install <local-path-or-name@marketplace>
copilot plugin update <plugin-name@marketplace-name>
copilot plugin uninstall <plugin-name>
```

3. Confirm MCP servers match the Copilot CLI config:

```bash
copilot mcp list
```

The task pane no longer imports agent/skill ZIP files or owns a separate MCP registry. Agents, skills, prompts, MCP servers, and plugin updates are owned by the Copilot CLI. On startup, the proxy automatically registers the Office Coding Agent marketplace when missing and ensures the required `office-excel`, `office-powerpoint`, `office-word`, and `office-outlook` plugins are installed and updated.

## Troubleshooting

- **Add-in not visible in Shared Folder**
  - Ensure Excel was restarted after `sideload:share:trust`.
  - Confirm the file exists in `%USERPROFILE%\OfficeAddinCatalog`.
- **Task pane doesn’t load on another machine**
  - The manifest probably points to `localhost`. Use `manifest.staging.xml`.
- **Share setup fails**
  - If script reports elevation required, open PowerShell as Administrator and rerun `npm run sideload:share:setup`.
- **Share cleanup fails**
  - If script reports elevation required, open PowerShell as Administrator and rerun `npm run sideload:share:cleanup`.
