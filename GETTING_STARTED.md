# Getting Started

Run Office Coding Agent locally — no installers required.

> **📖 See the [README](README.md) for architecture details, available scripts, and testing docs.**

## Prerequisites

| Software                        | Notes                                                      | Download                                                           |
| ------------------------------- | ---------------------------------------------------------- | ------------------------------------------------------------------ |
| **Node.js 20+**                 | Required to run the proxy server and build the add-in      | [nodejs.org](https://nodejs.org/)                                  |
| **Git**                         | Required to clone the repo                                 | [git-scm.com](https://git-scm.com/downloads)                       |
| **GitHub CLI**                  | Required for Copilot authentication                        | [cli.github.com](https://cli.github.com/)                          |
| **GitHub Copilot subscription** | Individual, Business, or Enterprise                        | [github.com/features/copilot](https://github.com/features/copilot) |
| **Microsoft Office**            | Excel, PowerPoint, or Word (Microsoft 365 or Office 2019+) | —                                                                  |

---

## Setup

### 1. Clone and install dependencies

```bash
git clone https://github.com/sbroenne/office-coding-agent.git
cd office-coding-agent
npm install
```

---

### 2. Authenticate with GitHub Copilot

Sign in with your GitHub account. The proxy server uses the GitHub CLI to manage Copilot authentication — no API keys or endpoint config needed.

```bash
gh auth login
```

Follow the browser prompt to complete sign-in. You only need to do this once.

> If you already use `gh` and are signed in, you can verify with `gh auth status`.

---

### 3. Register the add-in

This step trusts the local SSL certificate and registers the add-in manifest with Office so it appears in **My Add-ins**.

**Windows (run from a normal PowerShell — no elevation needed):**

```powershell
npm run register:win
```

**macOS:**

```bash
npm run register:mac
```

> Close and **fully quit** the Office application before registering, then reopen it afterwards. Office caches add-in registrations at startup.

---

### 4. Start the dev server and sideload

The proxy server bridges the browser task pane to the GitHub Copilot API via WebSocket. It must be running whenever you use the add-in.

```bash
npm run start:dev:desktop
```

This starts the dev server, waits for `https://localhost:3000`, then sideloads the add-in into Excel Desktop. Use a host-specific variant for other Office apps:

```bash
npm run start:dev:desktop:excel
npm run start:dev:desktop:ppt
npm run start:dev:desktop:word
npm run start:dev:desktop:outlook
```

The server handles both the Vite dev server (task pane UI) and the Copilot WebSocket proxy on port 3000.

---

### 5. Sideload into Office without starting the server

If the dev server is already running, you can sideload only:

```bash
# Excel
npm run start:desktop:excel

# PowerPoint
npm run start:desktop:ppt

# Word
npm run start:desktop:word
```

This opens the Office application and injects the add-in. The task pane will appear automatically.

> **Alternative — use My Add-ins:** Once registered (step 3), you can also open the add-in manually from Office via **Insert → Add-ins → My Add-ins → Office Coding Agent**, without running the sideload command each time. The proxy server (step 4) must still be running.

---

### 6. Start chatting

The task pane opens with an AI chat interface. Type a message to get started.

- Use the **Model picker** (bottom of the input bar) to choose a Copilot model.
- Type `/skill-name` or `/prompt-name` to invoke installed Copilot CLI skills and `.prompt.md` prompt files.
- Use the **MCP servers** picker (bottom of the input bar) to enable, disable, and sign in to MCP servers from `copilot mcp list`.
- Use `copilot plugin` CLI commands to install, update, or remove user plugins. On startup, the proxy automatically ensures the Office Coding Agent marketplace plugins (`office-excel`, `office-powerpoint`, `office-word`, `office-outlook`) are installed and updated in your normal Copilot CLI config.
- Use the **New Conversation** button (header) to reset the chat.

---

## Stopping

To stop the sideload session:

```bash
npm run stop
```

To stop the proxy server, close the dev server window or run:

```bash
npm run dev:stop
```

---

## Uninstalling

**Windows:**

```powershell
npm run unregister:win
```

**macOS:**

```bash
npm run unregister:mac
```

This removes the manifest registration and cleans up the trusted certificate entry. Fully quit and reopen Office after unregistering.

---

## Troubleshooting

### Add-in not appearing in My Add-ins

- Make sure you fully quit and restarted Office **after** running `npm run register:win/mac`.
- Confirm the proxy server is running on `https://localhost:3000`.
- Re-run the register script.

### Task pane loads but shows a connection error

- The proxy server is not running. Start it with `npm run dev`.
- Check that port 3000 is not blocked by another process: `npm run stop` and retry.

### SSL certificate errors in Office

- Re-run `npm run register:win` (Windows) or `npm run register:mac` (macOS).
- On macOS you may be prompted for your password to trust the certificate in Keychain.

### "Not authenticated" or Copilot errors

- Run `gh auth status` to confirm you are signed in.
- Run `gh auth login` to re-authenticate if needed.

### MCP servers do not appear as expected

- Run `copilot mcp list` to see the source of truth used by the add-in.
- Add or remove servers with `copilot mcp add` and `copilot mcp remove`.
- Reload the task pane or start a new conversation after changing MCP config so the next session uses the updated CLI list.
- If a remote MCP server needs sign-in, use the **MCP servers** picker action or follow the foreground sign-in prompt.

### Tray mode (alternative to `npm run dev`)

If you prefer a system tray app instead of a terminal:

```bash
npm run start:tray
```

Then sideload from the tray menu or use:

```bash
npm run start:tray:excel
npm run start:tray:ppt
npm run start:tray:word
```
