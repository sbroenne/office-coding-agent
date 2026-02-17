# Visualization Workflow

Use when creating or refining charts and visual summaries in Excel.

## Workflow

1. Confirm chart intent (trend, comparison, composition, distribution).
2. Ensure source data is clean and correctly typed.
3. Create chart with an appropriate default type.
4. Improve readability (title, labels, axis meaning, layout).
5. Verify chart reflects the intended range and interpretation.

## Rules

- Prefer simple chart types unless complexity is explicitly requested.
- Add explicit titles that describe the metric and scope.
- Avoid overcrowded visuals; keep categories and series legible.
- Verify the chart source range after data updates.

## Tool Patterns

- Prepare data: `get_range_values` + `set_number_format`
- Create chart: `create_chart`
- Tune semantics: `set_chart_type`, `set_chart_title`
- Final placement/readability: keep chart near source data or on a report sheet
