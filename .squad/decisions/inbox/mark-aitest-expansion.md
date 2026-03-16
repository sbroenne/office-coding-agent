# AI test manifest strategy

## Context
`tests-aitest/` now covers Excel, PowerPoint, Word, and Outlook using `pytest-skill-engineering` with manifest-driven MCP simulators.

## Decision
Use per-host manifest files under `tests-aitest/manifests/` instead of relying on a single shared manifest path.

- Excel manifest is generated from the decomposed config arrays in `src/tools/configs/` because the legacy Excel simulator and tests assert against decomposed tool names like `set_range_values` and `add_color_scale`.
- PowerPoint, Word, and Outlook manifests are generated from the runtime `Tool[]` objects exposed by `src/tools/` because those hosts either already use runtime tools directly or do not have a complete config-array source equivalent.

## Why
This keeps the eval schema aligned with the real host surfaces while preserving existing Excel simulator routing and enabling new host suites without writing separate handwritten manifests.

## Follow-up
If the project later standardizes all hosts on a single manifest shape, revisit the Excel simulator and tests so they can move off the decomposed config manifest and onto the runtime `Tool[]` surface too.
