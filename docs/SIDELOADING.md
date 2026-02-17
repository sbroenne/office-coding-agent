# Sideloading Guide

This project supports three sideloading lanes:

1. **Local desktop dev (fastest)** via `localhost`
2. **Local shared folder catalog** (Windows testing flow)
3. **Staging manifest** that points to **GitHub Pages**

## Important Model

A shared folder catalog distributes the **manifest only**.

The add-in web app (task pane HTML/JS/CSS) must be hosted at the HTTPS URLs in the manifest (`SourceLocation`, icon URLs).

- For one-machine local dev, `https://localhost:3000` is fine.
- For testing from other machines, use `manifests/manifest.staging.xml` that points to GitHub Pages.

## Lane 1: Local Desktop Dev

```bash
npm run start:desktop
```

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
npm run manifest:staging
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

Generated file:

- `manifests/manifest.staging.xml`

Regenerate anytime:

```bash
npm run manifest:staging
```

## Import Checklist (Skills & Agents)

Use this after the add-in is loaded (desktop or staging) to verify ZIP import flows.

1. Generate sample ZIPs

```bash
npm run extensions:samples
```

Expected files:

- `samples/extensions/sample-skills.zip`
- `samples/extensions/sample-agents.zip`

2. Open the task pane and open picker management

- In Excel, open the add-in task pane.
- For agents: open **Agent picker** in the input toolbar, then click **Manage agents…**
- For skills: open **Skill picker** in the header, then click **Manage skills…**

3. Import agents ZIP

- Click **Import Agents ZIP**.
- Select `sample-agents.zip`.
- Verify success message appears.
- Verify imported agents appear under **Imported** list.
- Verify bundled agents remain under **Bundled (read-only)**.

4. Import skills ZIP

- Click **Import Skills ZIP**.
- Select `sample-skills.zip`.
- Verify success message appears.
- Verify imported skills appear under **Imported** list.
- Verify bundled skills remain under **Bundled (read-only)**.

5. Verify pickers

- Agent picker and Skill picker should show grouped sections:
  - **Bundled**
  - **Imported**

6. Verify remove behavior

- Remove one imported agent and one imported skill from Settings.
- Confirm they disappear from Imported lists and pickers.
- Confirm bundled entries cannot be removed.

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
