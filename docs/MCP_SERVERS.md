# MCP Servers

Office Coding Agent keeps MCP handling lightweight. The task pane can toggle the built-in MCP servers that are explicitly bundled with the add-in, while plugin-provided MCP servers remain owned by the Copilot CLI/SDK.

## Built-in Servers

Built-in MCP servers are defined in `BUNDLED_MCP_SERVERS` (`src/types/settings.ts`) and exposed through the MCP picker in the chat toolbar.

| Name | Transport | Description |
|---|---|---|
| `workiq` | stdio (`npx -y @microsoft/workiq mcp`) | Microsoft 365 Copilot — emails, meetings, documents, Teams |
| `powerbi` | HTTP (`https://api.fabric.microsoft.com/v1/mcp/powerbi`) | Power BI — query semantic models, generate DAX, explore data |

The picker only enables/disables these known servers for the current chat session. The app no longer has an MCP server management dialog, import flow, log viewer, or app-owned MCP registry.

## Plugin MCP Servers

Copilot CLI plugins may include MCP server configuration, but Office Coding Agent does not parse plugin directories or merge plugin MCP servers itself. Install, update, and remove plugins with the Copilot CLI:

```bash
copilot plugin list
copilot plugin install <source-or-name@marketplace>
copilot plugin update <plugin-name@marketplace-name>
copilot plugin uninstall <plugin-name>
```

See GitHub's Copilot CLI plugin docs for plugin and marketplace authoring:

- [About Copilot CLI plugins](https://docs.github.com/en/copilot/concepts/agents/copilot-cli/about-cli-plugins)
- [Customize Copilot CLI with plugins and marketplaces](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/plugins-marketplace)

## Relevant Files

| File | Purpose |
|---|---|
| `src/types/settings.ts` | Built-in MCP server definitions and disabled-server settings |
| `src/components/McpPicker.tsx` | Lightweight enable/disable picker |
| `src/services/mcp/mcpService.ts` | Converts built-in MCP server config to Copilot SDK config |
| `src/hooks/useOfficeChat.ts` | Passes enabled built-in MCP servers to session creation |
| `src/copilotProxy.mjs` | Forwards MCP lifecycle status/log notifications from the SDK |
