# Modeling Workflow

Use when requests involve formulas, assumptions, dependencies, scenario logic, or model refactoring.

## Principles

- Keep assumptions in explicit input cells/tables, separate from calculated outputs.
- Use predictable row/column structure so formulas can be filled or copied safely.
- Prefer readable formulas over compact but opaque expressions.
- Reuse shared assumptions to avoid drift across sheets.

## Workflow

1. Identify input ranges and expected output ranges before writing formulas.
2. Write formulas in the minimal target range first.
3. Re-read formulas to confirm references and relative/absolute behavior.
4. Trigger recalculation when model changes are complete.
5. Re-read key outputs and sanity-check totals/signs/order of magnitude.

## Quality Checks

- Validate that totals reconcile (subtotals roll up to grand totals).
- Spot-check boundary cases (zero, blank, negative, large values).
- Confirm copied formulas preserve intended references.
- Flag circular references or volatile formula overuse when detected.

## Tool Patterns

- Read model state: `get_range_values`, `get_range_formulas`
- Apply formulas: `set_range_formulas`
- Recalculate: `recalculate_workbook`
- Structure check: `analyze_data`
