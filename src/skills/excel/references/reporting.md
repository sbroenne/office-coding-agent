# Reporting Workflow

Use when building summaries, KPI blocks, rollups, and report-ready worksheet outputs.

## Workflow

1. Discover and read source data.
2. Build summary structures (table, pivot, or formula-driven section).
3. Apply presentation formatting and labels.
4. Validate totals and key metrics.
5. Present concise findings.

## Rules

- Keep source data and reporting output separated when possible.
- Use tables for structured reporting ranges.
- Ensure number formats match metric semantics (currency, percent, date, integer).
- Include high-signal metrics by default; avoid clutter.

## Tool Patterns

- Source discovery: `list_tables`, `get_used_range`
- Build report table: `set_range_values` -> `create_table`
- KPI formulas: `set_range_formulas` + `set_number_format`
- Pivots: `create_pivot_table` + `add_pivot_field`
- Final polish: `format_range`, `auto_fit_columns`
