# Changelog

## Unreleased

- Switched plugin handling to the Copilot CLI as the single source of truth.
- Added startup bootstrap that registers the Office Coding Agent marketplace, installs missing required Office plugins, and updates those plugins through the user's normal Copilot CLI configuration.
- Removed app-owned plugin discovery, plugin management routes/tools, plugin slash-command discovery, and MCP management/import UI.
