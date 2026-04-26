# Changelog

## Unreleased

- Switched plugin handling to the Copilot CLI as the single source of truth.
- Added startup bootstrap that registers the Office Coding Agent marketplace, installs missing required Office plugins, and updates those plugins through the user's normal Copilot CLI configuration.
- Removed app-owned plugin management routes/tools, plugin session injection, and MCP management/import UI.
- Added task-pane slash suggestions for installed CLI skills using the documented `/skill-name` invocation pattern.
