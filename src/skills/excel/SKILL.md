---
name: excel
description: General-purpose Excel skill for workbook analysis, transformation, reporting, and visualization tasks.
license: MIT
---

# Excel Default Skill

Use this as the default orchestration skill for Excel tasks.

## Operating Loop

1. **Discover** — Inspect workbook structure and existing data first.
2. **Read** — Read exact target ranges/tables before any mutation.
3. **Execute** — Apply focused updates with the narrowest possible write scope.
4. **Verify** — Re-read key outputs and confirm formulas/values are correct.
5. **Summarize** — Finish with a concise plain-language change summary.

## Delegated Guidance

Use the focused reference docs in `references/` when task depth requires it:

- Data quality workflow: `references/data-quality.md`
- Reporting workflow: `references/reporting.md`
- Visualization workflow: `references/visualization.md`
- Modeling workflow: `references/modeling.md`

## Always-On Defaults

- Prefer targeted updates over delete/rebuild operations.
- Always apply formats after writes when data type is known.
- Always finish with a clear summary of changes made.

## High-Level Tool Guidance

| Task                        | Primary Tool                                               |
| --------------------------- | ---------------------------------------------------------- |
| Discover workbook structure | `get_workbook_info`, `list_sheets`, `list_tables`          |
| Inspect data before changes | `get_used_range`, `get_range_values`, `get_range_formulas` |
| Write values/formulas       | `set_range_values`, `set_range_formulas`                   |
| Apply formatting            | `set_number_format`, `format_range`, `auto_fit_columns`    |
| Manage structured data      | `create_table`, `get_table_data`, `filter_table`           |
| Build visuals               | `create_chart`, `set_chart_type`, `set_chart_title`        |
| Build models/calculations   | `set_range_formulas`, `recalculate_workbook`, `analyze_data` |

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue independent remaining steps.
