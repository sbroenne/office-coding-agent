# MCP Servers

Office Coding Agent treats the Copilot CLI configuration and installed plugin manifests as the source of truth for MCP servers. The add-in does not ship a hardcoded MCP registry. The proxy combines the CLI list with installed plugin MCP definitions and resolves execution settings server-side; the browser receives display summaries only.

## Source of Truth

Built-in Office tools are separate from MCP servers. They are registered directly with the Copilot SDK, and their calls are forwarded to the task pane for Office.js execution. They do not appear as an Office MCP server in the picker.

The local proxy exposes `/api/mcp-servers` by running:

```bash
copilot mcp list --json
```

`useOfficeChat` fetches these display summaries and sends server selection settings to the proxy. The proxy resolves full server configuration when creating the SDK session. The MCP picker shows CLI- and plugin-configured servers, lets users enable/disable servers for the Office session, and surfaces sign-in/retry/switch-account actions for authenticated remote servers.

To change what appears in the add-in, update the Copilot CLI MCP config:

```bash
copilot mcp list
copilot mcp get <server-name>
copilot mcp add <server-name> <url-or-command-and-args>
copilot mcp remove <server-name>
```

## Plugin MCP Servers

Copilot CLI plugins may include MCP server configuration. Install, update, and remove plugins with the Copilot CLI:

Required Office plugins are installed and updated automatically before the local server accepts chats. This happens on each server startup, not continuously; restart the server to pick up merged plugin changes. Update failures are logged in its terminal.

```bash
copilot plugin list
copilot plugin install <source-or-name@marketplace>
copilot plugin update <plugin-name@marketplace-name>
copilot plugin uninstall <plugin-name>
```

See GitHub's Copilot CLI plugin docs for plugin and marketplace authoring:

- [About Copilot CLI plugins](https://docs.github.com/en/copilot/concepts/agents/copilot-cli/about-cli-plugins)
- [Customize Copilot CLI with plugins and marketplaces](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/plugins-marketplace)

## OAuth Recovery

Remote HTTP/SSE servers that require authentication use SDK-owned OAuth recovery. When sign-in is required, the task pane shows a foreground prompt and the MCP picker offers **Sign in**, **Retry sign in**, or **Switch account** actions. Login hints are normalized before the proxy starts OAuth.

## Relevant Files

| File                                  | Purpose                                                     |
| ------------------------------------- | ----------------------------------------------------------- |
| `src/plugins/cliMcpServers.mjs`       | Runs and normalizes `copilot mcp list --json`               |
| `src/server.mjs`                      | Serves `/api/mcp-servers` from the CLI config               |
| `src/services/mcp/mcpServerConfig.ts` | Browser helper for loading CLI MCP server config            |
| `src/components/McpPicker.tsx`        | Enables/disables CLI MCP servers and starts OAuth recovery  |
| `src/hooks/useOfficeChat.ts`          | Sends MCP server selection settings to the proxy            |
| `src/copilotProxy.mjs`                | Forwards MCP lifecycle and OAuth notifications from the SDK |
