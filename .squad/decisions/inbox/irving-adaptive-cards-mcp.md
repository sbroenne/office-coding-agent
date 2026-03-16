# Irving Decision Note — adaptive-cards-mcp integration

## Context
The project already supports stdio MCP servers, and the Copilot CLI MCP config is the quickest place to make a new capability available to agents during development. Adaptive Cards are a strong fit for Office-adjacent workflows because they turn structured content into compact, actionable UI payloads that can travel into Outlook, Teams, and other supported hosts.

## Decision
Add `adaptive-cards-mcp` to `.copilot/mcp-config.json` as a stdio MCP server launched through `npx -y adaptive-cards-mcp`, while keeping the existing `EXAMPLE-github` entry intact. Document the capability in a Squad skill so agents know when and how to use the server for Office-oriented work.

## Available tools
The integration adds these seven Adaptive Card tools:
- `generate_card`
- `validate_card`
- `data_to_card`
- `optimize_card`
- `template_card`
- `transform_card`
- `suggest_layout`

## Why
This gives the team a lightweight way to generate, validate, and adapt Adaptive Card payloads without building custom card tooling into the repo. The biggest near-term wins are Outlook-oriented actionable-message concepts and Excel-to-card data visualization flows, with additional value when content needs to be reused in Teams or other supported collaboration surfaces.

## Most relevant Office hosts
- **Outlook** — best fit for concise, actionable card payloads and mail-centric workflows
- **Excel** — best fit for converting structured worksheet data into summary cards for downstream sharing
- **Teams integrations** — strong collaboration target when the same Office-derived content needs richer sharing surfaces

## Verification status
A direct verification attempt with `npx -y adaptive-cards-mcp --help` currently fails with an npm 404 for `adaptive-cards-mcp`, even though the GitHub repository documents that package name. The config and skill were still added as requested, but the package appears not to be reachable from the npm registry in this environment today; a follow-up may be needed to publish the package or point the config at a different distributable.

## Impacted paths
- `.copilot/mcp-config.json`
- `.squad/skills/adaptive-cards/SKILL.md`
- `.squad/decisions/inbox/irving-adaptive-cards-mcp.md`
