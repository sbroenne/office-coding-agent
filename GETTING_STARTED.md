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

### 4. Start the proxy server

The proxy server bridges the browser task pane to the GitHub Copilot API via WebSocket. It must be running whenever you use the add-in.

```bash
npm run dev
```

You should see output like:

```
  Copilot Office Add-in server running on https://localhost:3000
  API: https://localhost:3000/api
```

Leave this terminal open. The server handles both the Vite dev server (task pane UI) and the Copilot WebSocket proxy on port 3000.

---

### 5. Sideload into Office

Open a **second terminal** and run the sideload command for your target host:

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

- Use the **Agent picker** (bottom of the input bar) to switch agents.
- Use the **Model picker** (bottom of the input bar) to choose a Copilot model.
- Use the **Skill picker** (header icon) to toggle context skills on/off.
- Use the **New Conversation** button (header) to reset the chat.

---

## Stopping

To stop the sideload session:

```bash
npm run stop
```

To stop the proxy server, press `Ctrl+C` in the terminal where `npm run dev` is running.

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
