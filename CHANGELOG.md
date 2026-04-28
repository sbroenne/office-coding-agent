# Changelog

## Unreleased

- Switched plugin handling to the Copilot CLI as the single source of truth.
- Added startup bootstrap that registers the Office Coding Agent marketplace, installs missing required Office plugins, and updates those plugins through the user's normal Copilot CLI configuration.
- Removed app-owned plugin management routes/tools, plugin session injection, and MCP management/import UI.
- Added task-pane slash suggestions for installed CLI skills and `.prompt.md` prompt files using the documented `/name` invocation pattern.
- Added one-command local desktop startup via `npm run start:dev:desktop`, which starts the dev server and then sideloads the add-in.
- Added authenticated remote MCP server recovery with SDK-owned OAuth, login-hint support, foreground sign-in prompts, and MCP picker Sign in/Retry/Switch account actions.
- Switched MCP server discovery and session wiring to `copilot mcp list --json` so the picker and active Copilot session match the user's CLI MCP config.
- Removed app-owned agent files, local agent frontmatter parsing, agent ZIP import/export helpers, and browser-provided `customAgents`; the Agent picker now selects CLI-owned agents installed by the Copilot CLI plugins.
