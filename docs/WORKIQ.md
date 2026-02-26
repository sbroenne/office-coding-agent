# WorkIQ Integration

WorkIQ brings **Microsoft 365 Copilot** data (emails, meetings, documents, Teams messages, calendar) into the office-coding-agent chat. When enabled, the AI assistant can query your M365 data to answer workplace questions.

## Prerequisites

| Requirement | Details |
|---|---|
| **Node.js** | ≥ 18 (for `npx`) |
| **npm package** | `@microsoft/workiq` — installed on-the-fly via `npx -y @microsoft/workiq mcp` |
| **Microsoft 365 account** | Work/school account with Copilot license |
| **Authentication** | WorkIQ uses device-code / browser-based OAuth via Microsoft Entra ID. On first use it opens a browser window for consent. |
| **Network** | Access to `graph.microsoft.com` and Entra ID endpoints |

## How It Works

```
Browser task pane (React + assistant-ui)
         ↓ WebSocket
Node.js proxy server (src/server.mjs)
         ↓ @github/copilot-sdk (tool routing)
         ├─ Office tools (Excel/Word/PowerPoint)
         └─ WorkIQ MCP server (stdio)
              ↓ Microsoft Graph API
         Microsoft 365 data
```

WorkIQ runs as an **MCP (Model Context Protocol) server** in `stdio` transport mode. The Copilot SDK spawns it as a child process when enabled. The MCP server exposes tools that the AI model can call to query M365 data.

## Configuration

### MCP Server Definition

The WorkIQ MCP server is defined as a built-in constant in `src/types/settings.ts`:

```typescript
export const WORKIQ_MCP_SERVER: McpServerConfig = {
  name: 'workiq',
  description: 'Microsoft 365 Copilot — emails, meetings, documents, Teams',
  transport: 'stdio',
  command: 'npx',
  args: ['-y', '@microsoft/workiq', 'mcp'],
};
```

A matching `workiq-mcp.json` exists at the repo root for reference:

```json
{
  "mcpServers": {
    "workiq": {
      "command": "npx",
      "args": ["-y", "@microsoft/workiq", "mcp"]
    }
  }
}
```

### User Settings

| Setting | Type | Default | Description |
|---|---|---|---|
| `workiqEnabled` | `boolean` | `false` | Master toggle for WorkIQ |
| `workiqModel` | `string \| null` | `null` | Optional model override. When set, the entire Copilot session uses this model instead of `activeModel`. When `null`, falls back to the main model. |

Both settings are persisted via `OfficeRuntime.storage` (Zustand store with `officeStorage` adapter).

## Enabling WorkIQ

### Via the UI

1. Open the add-in task pane
2. In the **ChatHeader**, click the **WorkIQ** toggle button
3. When enabled:
   - A green status dot appears on the toggle
   - A small **Model** dropdown appears next to the toggle (optional — pick a different model for WorkIQ sessions)
4. The chat session automatically resets to apply the change

### Via Code

```typescript
import { useSettingsStore } from '@/stores/settingsStore';

// Enable WorkIQ
useSettingsStore.getState().toggleWorkiq();

// Set a specific model for WorkIQ sessions
useSettingsStore.getState().setWorkiqModel('gpt-4.1');
```

## Session Behavior

When WorkIQ is enabled:

1. **`useOfficeChat`** reads `workiqEnabled` from the settings store
2. The WorkIQ MCP server (`WORKIQ_MCP_SERVER`) is added to the MCP server list passed to the Copilot SDK
3. If `workiqModel` is set, it overrides `activeModel` for the entire session
4. **Important:** Toggling WorkIQ on/off triggers a full session reset (new WebSocket + new Copilot session). Chat history is lost.

The MCP server is filtered out of the regular `importedMcpServers` list — it's managed separately via the dedicated toggle.

## What WorkIQ Can Do

Once enabled, the AI assistant gains access to M365 tools:

- **Emails** — search, read, summarize emails from Outlook
- **Meetings** — find upcoming meetings, read agendas, attendees
- **Documents** — search SharePoint/OneDrive files
- **Teams** — read Teams messages and channels
- **Calendar** — check availability, find events
- **People** — look up colleagues, org structure

Example prompts:
- "What emails did I get from Sarah today?"
- "What's on my calendar this week?"
- "Summarize the last Teams discussion about the Q4 budget"
- "Find documents about the project proposal"

## EULA

WorkIQ requires accepting a EULA on first use. The EULA URL is: https://github.com/microsoft/work-iq-mcp

## Troubleshooting

| Problem | Solution |
|---|---|
| WorkIQ tools not appearing | Check that the toggle is enabled (green dot visible). Check browser console for MCP spawn errors. |
| Authentication fails | Ensure you have a valid M365 work/school account. Try clearing browser auth cache. Check Conditional Access policies. |
| "npx not found" | Ensure Node.js ≥ 18 is installed and `npx` is on PATH. |
| Session resets on toggle | This is expected — enabling/disabling WorkIQ creates a new Copilot session. |
| Model dropdown not showing | The dropdown only appears when WorkIQ is enabled AND model list has been fetched from the Copilot API. |

## Key Files

| File | Role |
|---|---|
| `src/types/settings.ts` | `WORKIQ_MCP_SERVER` constant, `workiqEnabled`/`workiqModel` in `UserSettings` |
| `src/stores/settingsStore.ts` | `toggleWorkiq()`, `setWorkiqModel()` actions |
| `src/hooks/useOfficeChat.ts` | Injects WorkIQ MCP server into session, applies model override |
| `src/components/ChatHeader.tsx` | UI toggle + model dropdown |
| `workiq-mcp.json` | Standalone MCP config (reference / external tooling) |
