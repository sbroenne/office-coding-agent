# Squad Decisions

## Active Decisions

### User Directive: Python Testing Framework & Expansion (2026-03-15T10:27Z)
**Owner:** Stefan Broenner (via Copilot)  
**Status:** Captured  
Keep the Python test project (pyproject.toml + tests-aitest/). Switch to `pytest-skill-engineering` as the testing framework. Add tests for ALL tools (PowerPoint, Word, Outlook, not just Excel). Document in README.

### Harmony: Squad v0.8.25 Reference Templates (2026-03-15)
**Owner:** Harmony  
**Status:** Completed  
Created missing v0.8.25 reference templates and plugin marketplace state file:
- `.squad/templates/ralph-reference.md`, `casting-reference.md`, `copilot-agent.md`, `human-members.md`, `issue-lifecycle.md`, `prd-intake.md`, `ceremony-reference.md`
- `.squad/plugins/marketplaces.json`

**Why:** The coordinator governance file (`.github/agents/squad.agent.md`) referenced these files as on-demand detail expansions but they did not exist. Restores lookup paths and ensures plugin marketplace has a defined empty-state file.

### Irving: Adaptive Cards MCP Integration (2026-03-15)
**Owner:** Irving  
**Status:** Completed  
Added `adaptive-cards-mcp` to `.copilot/mcp-config.json` as a stdio MCP server via `npx -y adaptive-cards-mcp`. Exposes seven tools: `generate_card`, `validate_card`, `data_to_card`, `optimize_card`, `template_card`, `transform_card`, `suggest_layout`.

**Why:** Lightweight way to generate and validate Adaptive Card payloads for Office workflows. Best fit for Outlook actionable messages and Excel-to-card data visualization flows.

**Note:** npm 404 for package on current environment; follow-up may be needed to publish or point config at different distributable.

### Irving: Upgrade @github/copilot to 1.0.5 (2026-03-15)
**Owner:** Irving  
**Status:** Completed  
Upgraded critical security packages:
- `@github/copilot`: ^0.0.414 → ^1.0.5 (major version bump)
- `@github/copilot-sdk`: ^0.1.30 → ^0.1.32

**Why:** GHSA-g8r9-g2v8-jv6f vulnerability (arbitrary code execution via shell expansion). No breaking API changes affecting codebase. Build, typecheck, and integration tests all pass.

**Impact:** None required beyond package.json/package-lock.json. Fully backward compatible.

### Irving: ESLint 10 Vitest Plugin (2026-03-15)
**Owner:** Irving  
**Status:** Completed  
Switched from `eslint-plugin-vitest` (fails under ESLint 10) to `@vitest/eslint-plugin` in `eslint.config.mjs`.

**Why:** `eslint-plugin-vitest` 0.5.4 only supports ESLint 8/9. `@vitest/eslint-plugin` declares support for ESLint >=8.57.0, covers ESLint 10, preserves flat-config usage.

**Future:** Treat `@vitest/eslint-plugin` as the supported Vitest ESLint integration going forward.

### Irving: Manifest-Backed Excel MCP Descriptions (2026-03-15)
**Owner:** Irving  
**Status:** Proposed  
Use `tests-aitest/manifests/excel-tools-manifest.json` as the source of truth for Excel AI eval tool descriptions. `tests-aitest/excel_mcp.py` maps decomposed tool names back to aggregate manifest tools and reuses parameter descriptions.

**Why:** Keeps eval MCP aligned with product tool definitions. Avoids second, lower-quality description system in Python. Makes decomposed tools self-describing by attaching the conditional-format or validation type.

### Irving: Standardize Python Evals on pytest-skill-engineering (2026-03-15)
**Owner:** Irving  
**Status:** Proposed  
Move `tests-aitest/` from `pytest-aitest` to `pytest-skill-engineering`. Remove `pytest-llm-assert` dependency (now built in).

**Why:** `pytest-skill-engineering` is current framework for MCP/prompt/skill evals. Preserves MCPServer/Provider/Wait concepts; renames Agent→Eval and aitest_run→eval_run. Consolidates Python eval dependencies.

### Mark: Adversarial Excel AI Evals (2026-03-15)
**Owner:** Mark  
**Status:** Proposed  
Added `tests-aitest/test_excel_adversarial.py` with `pytest.mark.adversarial` marker. Decision: keep adversarial evals strict when wrong tool materially changes workbook behavior.

**Why:** Live execution showed prompt "Highlight values above 100 but don't restrict input" still called both `add_cell_value_format` AND `set_number_validation`, changing user behavior in the sheet. Real product-quality issue, not just prompt quirk.

**Recommendation:** Treat confusion tests as behavior tests. Assert correct tool called AND wrong tool not called when side effects differ. Keep `allowed_tools` narrow to identify genuine routing problems.

### Mark: AI Test Manifest Strategy (2026-03-15)
**Owner:** Mark  
**Status:** Proposed  
Use per-host manifest files under `tests-aitest/manifests/` instead of single shared manifest:
- **Excel:** Generated from decomposed config arrays in `src/tools/configs/` (aligns with legacy simulator and test expectations for names like `set_range_values`, `add_color_scale`)
- **PowerPoint/Word/Outlook:** Generated from runtime `Tool[]` objects (no config-array equivalent)

**Why:** Keeps eval schema aligned with real host surfaces. Preserves existing Excel simulator routing. Enables new host suites without handwritten manifests.

**Follow-up:** If project standardizes all hosts on single manifest shape, revisit Excel simulator to move off decomposed config manifest onto runtime `Tool[]` surface.

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction
