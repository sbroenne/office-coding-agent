# Harmony Decision Note — Squad v0.8.25 reference templates

## Context
The coordinator governance file (`.github/agents/squad.agent.md`) was already upgraded to v0.8.25 and referenced seven on-demand reference templates that did not exist in `.squad/templates/`. When the coordinator tried to load those references, the lookup failed silently, which degraded Ralph, casting, issue lifecycle, PRD intake, ceremony, human-member, and `@copilot` guidance.

## Decision
Create the missing v0.8.25 reference templates and the empty plugin marketplace state file expected by the plugin marketplace flow.

Created:
- `.squad/templates/ralph-reference.md`
- `.squad/templates/casting-reference.md`
- `.squad/templates/copilot-agent.md`
- `.squad/templates/human-members.md`
- `.squad/templates/issue-lifecycle.md`
- `.squad/templates/prd-intake.md`
- `.squad/templates/ceremony-reference.md`
- `.squad/plugins/marketplaces.json`

## Why
The governance file already depends on these references as on-demand detail expansions. Adding them restores the coordinator's documented lookup paths, gives Squad operators concrete formats and examples for the new v0.8.25 flows, and ensures plugin marketplace lookup has a defined empty-state file instead of a missing-path failure.

## Impacted paths
- `.squad/templates/ralph-reference.md`
- `.squad/templates/casting-reference.md`
- `.squad/templates/copilot-agent.md`
- `.squad/templates/human-members.md`
- `.squad/templates/issue-lifecycle.md`
- `.squad/templates/prd-intake.md`
- `.squad/templates/ceremony-reference.md`
- `.squad/plugins/marketplaces.json`
