# Irving Decision Note — Manifest-backed Excel MCP descriptions

## Context
The Excel AI eval MCP server exposes 83 decomposed tools, but the generated manifest only contains 10 aggregate tool groups with the rich copy authored in the TypeScript tool configs. Auto-humanizing Python method names produced weak descriptions that hurt tool selection quality in evals.

## Decision
Use `tests-aitest/manifests/excel-tools-manifest.json` as the source of truth for Excel AI eval tool and parameter descriptions. `tests-aitest/excel_mcp.py` should map each decomposed tool name back to its aggregate manifest tool + action, then reuse manifest parameter descriptions with explicit aliases where decomposed names differ.

## Why
This keeps the eval MCP server aligned with the product tool definitions and avoids a second, lower-quality description system in Python. It also makes decomposed tools like `add_color_scale` and `set_list_validation` self-describing by attaching the fixed conditional-format or validation type implied by the tool name.

## Impacted paths
- `tests-aitest/excel_mcp.py`
- `tests-aitest/manifests/excel-tools-manifest.json`
- `tests-aitest/conftest.py`
