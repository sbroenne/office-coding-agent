# Data Quality Workflow

Use when cleaning, normalizing, validating, or repairing workbook data.

## Workflow

1. Profile source ranges/tables and identify blanks, type mismatches, and duplicates.
2. Normalize values, date formats, and casing with targeted updates.
3. Validate with explicit rules where user input is expected.
4. Verify by re-reading key ranges.

## Rules

- Never overwrite unknown formulas without reading them first.
- Prefer targeted cell/range updates over full rewrites.
- Preserve headers and table structure unless explicitly asked to redesign.
- If values are ambiguous, apply the least destructive transformation.

## Tool Patterns

- Profile values: `get_used_range` -> `get_range_values`
- Inspect formulas: `get_range_formulas`
- Normalize values: `set_range_values`
- Add constraints: `set_list_validation`, `set_number_validation`
- Verify results: `get_range_values`
