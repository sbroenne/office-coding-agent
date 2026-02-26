# MCP Server Management

The office-coding-agent supports **MCP (Model Context Protocol) servers** that provide additional AI tools. MCP servers can be built-in (bundled with the add-in) or user-imported.

## Overview

```
Browser task pane (React + assistant-ui)
         ↓ WebSocket
Node.js proxy server (src/server.mjs)
         ↓ @github/copilot-sdk (tool routing)
         ├─ Office tools (Excel/Word/PowerPoint)
         └─ MCP servers (stdio / http / sse)
              ↓ external APIs
         External data (M365, databases, etc.)
```

MCP servers run alongside Office tools. The Copilot SDK manages their lifecycle — spawning stdio processes, connecting to HTTP/SSE endpoints, and routing tool calls.

## Bundled Servers

The add-in ships with built-in MCP servers defined in `BUNDLED_MCP_SERVERS` (`src/types/settings.ts`):

| Name | Transport | Description |
|---|---|---|
| `workiq` | stdio (`npx -y @microsoft/workiq mcp`) | Microsoft 365 Copilot — emails, meetings, documents, Teams |

Bundled servers:
- Appear in the MCP Manager dialog with a **"Built-in"** badge
- Can be **toggled on/off** but **not removed** or edited
- Are **disabled by default** — users must explicitly enable them
- Require explicit opt-in: when `activeMcpServerNames` is `null` (all servers active), only imported servers are included. Bundled servers must be explicitly listed.

## Adding MCP Servers

### Via the UI (MCP Manager Dialog)

1. Open **Settings** → **MCP Servers** (or the MCP button in the chat header)
2. Click **+ Add** to open the add server form
3. Enter:
   - **Name** — unique identifier
   - **Description** — optional
   - **Transport** — `stdio`, `http`, or `sse`
   - Transport-specific fields (command/args for stdio, URL/headers for http/sse)
4. Click **Add Server**

### Via mcp.json Import

1. Click **Import** in the MCP Manager dialog
2. Select a `mcp.json` file — supports both VS Code format (`servers` key) and Claude Desktop format (`mcpServers` key)

Example `mcp.json`:

```json
{
  "mcpServers": {
    "my-server": {
      "command": "npx",
      "args": ["-y", "my-mcp-server"]
    },
    "api-server": {
      "url": "https://api.example.com/mcp",
      "type": "http"
    }
  }
}
```

### Via Code

```typescript
import { useSettingsStore } from '@/stores/settingsStore';

useSettingsStore.getState().importMcpServers([
  {
    name: 'my-server',
    transport: 'stdio',
    command: 'npx',
    args: ['-y', 'my-mcp-server'],
  },
]);
```

## Managing Servers

The MCP Manager dialog provides VS Code-style server management:

### Status Indicators

Each server shows a colored status dot:
- 🟢 **Connected** — server is running and tools are available
- 🟡 **Starting** — server is being initialized
- 🔴 **Error** — server failed to connect (error message shown)
- ⚫ **Stopped** — server was active but is now stopped
- ⚪ **Disabled** — server is toggled off

### Actions

| Action | Button | Description |
|---|---|---|
| Start/Stop | ▶/■ | Toggle server enabled/disabled |
| Restart | ⟳ | Quick restart (toggle off then back on) |
| Show Output | 📄 | Open the log viewer for this server |
| Edit | ✏️ | Edit server config (imported servers only) |
| Remove | 🗑️ | Remove server (imported servers only) |

### Output Log Window

The bottom of the MCP Manager shows a **log viewer** for inspecting server output:
- **Per-server logs** — select a server to view its output
- **Timestamped entries** with level coloring (info=gray, warn=amber, error=red)
- **Copy** — copy all log content to clipboard
- **Clear** — clear logs for the selected server
- **Auto-scroll** — automatically scrolls to latest entries (manual scroll lock when you scroll up)

### Tool Discovery

When a server reports its available tools, they appear as an expandable list in the server row. Click the tool count to see tool names and descriptions.

## Architecture

### Settings Store

MCP server state is managed in `useSettingsStore` (persisted via OfficeRuntime.storage):

- `importedMcpServers` — user-imported server configs
- `activeMcpServerNames` — which servers are enabled (`null` = all imported servers active; bundled servers require explicit listing)
- `importMcpServers(configs)` — add new servers
- `removeMcpServer(name)` — remove a server
- `toggleMcpServer(name)` — toggle server enabled/disabled
- `updateMcpServer(name, config)` — update server config

### MCP Status Store

Runtime state is tracked in `useMcpStatusStore` (ephemeral, not persisted):

- Per-server: status, error, tools, logs
- Actions: `setStatus()`, `addLog()`, `setTools()`, `clearLogs()`, `clearAll()`
- Log buffer capped at 500 entries per server

### Proxy Notifications

The Node.js proxy (`copilotProxy.mjs`) emits JSON-RPC notifications for MCP lifecycle events:

- `mcp.status` — server status changes (starting, connected, error, stopped)
- `mcp.log` — timestamped log entries per server
- `mcp.tools` — discovered tools per server

These are forwarded to the browser via WebSocket and stored in `mcpStatusStore`.

### Session Behavior

- Changing MCP server config (add/remove/toggle) triggers a full session reinit
- The new session is created with the updated server list
- Chat history is lost on reinit (expected behavior)
- Bundled servers only participate when explicitly enabled

## WorkIQ (Built-in MCP Server)

WorkIQ provides **Microsoft 365 Copilot** data access:

### Prerequisites

| Requirement | Details |
|---|---|
| **Node.js** | ≥ 18 (for `npx`) |
| **npm package** | `@microsoft/workiq` — installed on-the-fly via `npx -y @microsoft/workiq mcp` |
| **Microsoft 365 account** | Work/school account with Copilot license |
| **Authentication** | Device-code / browser-based OAuth via Microsoft Entra ID |

### Capabilities

- **Emails** — search, read, summarize from Outlook
- **Meetings** — find upcoming meetings, read agendas, attendees
- **Documents** — search SharePoint/OneDrive files
- **Teams** — read Teams messages and channels
- **Calendar** — check availability, find events
- **People** — look up colleagues, org structure

### EULA

WorkIQ requires accepting a EULA on first use: https://github.com/microsoft/work-iq-mcp

## Troubleshooting

| Problem | Solution |
|---|---|
| MCP tools not appearing | Check that the server is enabled (green status dot). Check the log viewer for errors. |
| WorkIQ auth fails | Ensure you have a valid M365 work/school account. Try clearing browser auth cache. |
| "npx not found" | Ensure Node.js ≥ 18 is installed and `npx` is on PATH. |
| Session resets on server toggle | Expected — enabling/disabling servers creates a new Copilot session. |
| Server shows "Error" status | Check the log viewer for the specific error message. Common issues: network, auth, missing dependencies. |

## Key Files

| File | Role |
|---|---|
| `src/types/settings.ts` | `BUNDLED_MCP_SERVERS` array, `McpServerConfig` in `UserSettings` |
| `src/types/mcp.ts` | `McpServerStatus`, `McpLogEntry`, `McpServerState` types |
| `src/stores/settingsStore.ts` | Server CRUD actions, toggle, active server tracking |
| `src/stores/mcpStatusStore.ts` | Ephemeral runtime state (status, logs, tools) |
| `src/hooks/useOfficeChat.ts` | Injects MCP servers into Copilot sessions, forwards status events |
| `src/components/McpManagerDialog.tsx` | VS Code-style management dialog |
| `src/components/McpLogViewer.tsx` | Inspectable log window component |
| `src/components/McpAddServerForm.tsx` | Add/edit server form |
| `src/services/mcp/mcpService.ts` | MCP config parsing, `getAllMcpServers()` helper |
| `src/copilotProxy.mjs` | MCP lifecycle notifications (status, log, tools) |
| `src/lib/websocket-client.ts` | MCP notification handling in browser |
