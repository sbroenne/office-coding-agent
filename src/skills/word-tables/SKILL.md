---
name: word-tables
description: >
  Specialized skill for creating, reading, and modifying tables in Word documents.
  Covers table creation, row/column management, cell formatting, and data entry.
version: 1.0.0
license: MIT
hosts: [word]
---

# Tables & Structured Data Skill

Activate this skill when working with tables, structured data, or tabular content in Word.

## Table Creation

Use `insert_table` with a 2D data array:
```
[
  ["Header 1", "Header 2", "Header 3"],
  ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
  ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
]
```

### Style Options
- `style`: Word table style name (e.g., `"Grid Table 4 - Accent 1"`)
- First row is treated as header row

## Reading Tables

1. `get_document_overview` → find how many tables exist
2. `get_table_data` with `tableIndex` (0-based) → read cell contents

## Modifying Tables

### Add Rows
`add_table_rows` with:
- `tableIndex` — which table (0-based)
- `rows` — 2D array of new row data
- `insertAtEnd` — true to append, false to insert at beginning

### Add Columns
`add_table_columns` with:
- `tableIndex` — which table
- `headers` — array of column headers
- `data` — array of arrays for column data

### Delete Rows
`delete_table_row` with:
- `tableIndex` — which table
- `rowIndex` — which row to delete (0-based)

### Update Cells
`set_table_cell_value` with:
- `tableIndex`, `rowIndex`, `columnIndex`
- `value` — new text
- `bold` — optional bold formatting
- `shading` — optional background color

## Workflow: Modify Existing Table

1. `get_document_overview` → find tables
2. `get_table_data(tableIndex)` → read current contents
3. Apply changes (add rows, update cells, etc.)
4. `get_table_data(tableIndex)` → verify result
5. Fix if needed

## Workflow: Create Data Table from Scratch

1. Plan table structure (columns, rows)
2. `insert_table` with header + data rows
3. `get_table_data(tableIndex)` → verify
4. Apply cell formatting if needed with `set_table_cell_value`
5. Final verification

## Tips

- **Always read before modifying** — understand the current table structure
- **Use descriptive headers** — clear column names
- **Keep cell content concise** — tables work best with short text
- **Verify after every change** — tables are easy to break
