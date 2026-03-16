# Decision Proposal: Standardize Python AI evals on pytest-skill-engineering

- **Date:** 2026-03-15
- **Requested by:** Stefan Broenner
- **Proposed by:** Irving
- **Status:** Proposed

## Decision
Move the `tests-aitest/` Python eval suite from `pytest-aitest` to `pytest-skill-engineering` and remove the standalone `pytest-llm-assert` dependency.

## Rationale
- `pytest-skill-engineering` is Stefan's current framework for MCP, prompt, and skill evals.
- The new framework preserves the existing MCPServer/Provider/Wait concepts while renaming the core eval API from `Agent` to `Eval` and `aitest_run` to `eval_run`.
- `llm_assert` is now built in, so carrying a separate `pytest-llm-assert` dependency is unnecessary.

## Impact
- Python eval dependencies are consolidated around the newer framework.
- Existing Excel eval scenarios keep their current structure with minimal code churn.
- Validation can continue with `uv lock` and `uv run pytest tests-aitest/ --collect-only` before any paid model execution.
