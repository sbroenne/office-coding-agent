---
name: Word Document Editing
description: Skill for editing and formatting Microsoft Word documents via the Office JS API.
license: MIT
hosts: [word]
---

# Word Document Editing Skill

Use this as the default orchestration skill for Word document tasks.

## Operating Loop

1. **Discover** — Call `get_document_overview` to understand heading hierarchy, paragraph count, and tables.
2. **Read** — Use `get_document_content`, `get_document_section`, `get_selection_text`, or `get_selection` to inspect the current content.
3. **Execute** — Create, modify, or format content using the appropriate tool.
4. **Verify** — Re-read modified content to confirm changes (e.g. `get_document_section` or `get_selection_text`).
5. **Summarize** — Finish with a concise plain-language summary of what was done.

## High-Level Tool Guidance

| Task                              | Primary Tool                  |
| --------------------------------- | ----------------------------- |
| Understand document structure     | `get_document_overview`       |
| Read full document as HTML        | `get_document_content`        |
| Read a specific section           | `get_document_section`        |
| Get selected text (plain)         | `get_selection_text`          |
| Get selected content (OOXML)      | `get_selection`               |
| Replace entire document           | `set_document_content`        |
| Insert HTML at selection          | `insert_content_at_selection` |
| Insert a paragraph                | `insert_paragraph`            |
| Insert a page/section break       | `insert_break`                |
| Find and replace text             | `find_and_replace`            |
| Insert a table                    | `insert_table`                |
| Insert a bulleted/numbered list   | `insert_list`                 |
| Insert an image                   | `insert_image`                |
| Apply font styling to selection   | `apply_style_to_selection`    |
| Apply named paragraph style       | `apply_paragraph_style`       |
| Set paragraph formatting          | `set_paragraph_format`        |
| Get document metadata             | `get_document_properties`     |
| Get comments                      | `get_comments`                |
| List content controls             | `get_content_controls`        |
| Insert text at bookmark           | `insert_text_at_bookmark`     |
| Read headers and footers          | `get_headers_footers`         |
| Set header or footer content      | `set_header_footer`           |
| Read table contents by index      | `get_table_data`              |
| Add rows to existing table        | `add_table_rows`              |
| Add columns to existing table     | `add_table_columns`           |
| Delete a table row                | `delete_table_row`            |
| Set table cell value/formatting   | `set_table_cell_value`        |
| Insert a hyperlink                | `insert_hyperlink`            |
| Insert a footnote                 | `insert_footnote`             |
| Insert an endnote                 | `insert_endnote`              |
| Get footnotes and endnotes        | `get_footnotes_endnotes`      |
| Delete selected content           | `delete_content`              |
| Wrap selection in content control | `insert_content_control`      |
| Search text and apply formatting  | `format_found_text`           |
| List document sections            | `get_sections`                |

## Choosing Between Tools

- **`set_document_content`** replaces the ENTIRE document body. Only use when starting fresh or regenerating full content.
- **`insert_content_at_selection`** inserts rich HTML at the cursor. Best for adding formatted content at a specific location.
- **`insert_paragraph`** is simpler — it appends or prepends a paragraph to the document body. Use for quick text additions.
- **`apply_style_to_selection`** changes font-level formatting (bold, italic, size, color). Use for inline text styling.
- **`apply_paragraph_style`** applies a named Word style (e.g. "Heading 1") to paragraphs. Use for structural formatting.
- **`set_paragraph_format`** sets paragraph-level properties (alignment, spacing, indent). Use for layout adjustments.
- **`insert_list`** is the easiest way to create bullet or numbered lists via HTML insertion.
- **`get_table_data`** reads an existing table's contents. Use before modifying table structure.
- **`add_table_rows`** / **`add_table_columns`** — add rows or columns to an existing table. Reference by `tableIndex` (0-based).
- **`delete_table_row`** — remove a specific row from a table.
- **`set_table_cell_value`** — update a single cell's text and optionally set shading/bold.
- **`insert_hyperlink`** — inserts a clickable link at the selection via HTML.
- **`insert_footnote`** / **`insert_endnote`** — insert notes at the selection (requires WordApi 1.5).
- **`format_found_text`** — search for text and apply font formatting (bold, color, etc.) to all matches.
- **`get_headers_footers`** / **`set_header_footer`** — read/write section headers and footers.
- **`insert_content_control`** — wrap the selection in a content control (useful for template fields).
- **`delete_content`** — delete the currently selected content.

## Iterative Refinement Workflow

Never treat a document change as done after a single pass. Always:

1. **Read** — Inspect the current state before making changes.
2. **Modify** — Apply the requested changes.
3. **Verify** — Re-read the modified content to confirm correctness.
4. **Refine** — If something is off, adjust and verify again.

### What to check during refinement:
- **Formatting** — Are styles, fonts, and spacing correct?
- **Completeness** — Did you include all requested content?
- **Consistency** — Does the new content match the document's existing style?
- **Structure** — Are headings, lists, and tables properly formed?

## Common Workflows

### Summarize a document
1. `get_document_overview` → understand structure
2. `get_document_content` → read full text
3. Provide a concise summary

### Add a new section
1. `get_document_overview` → understand current headings
2. `insert_paragraph` with style "Heading 1" → add heading
3. `insert_content_at_selection` → add section body content
4. Verify with `get_document_section`

### Format existing text
1. `get_selection_text` → see what's selected
2. `apply_style_to_selection` or `apply_paragraph_style` → apply formatting
3. `get_selection_text` → confirm result

### Create a table from data
1. `insert_table` with data array → insert structured table
2. `get_document_content` → verify table appears correctly

### Modify an existing table
1. `get_document_overview` → find how many tables exist
2. `get_table_data` with `tableIndex` → read current contents
3. `add_table_rows` / `add_table_columns` / `delete_table_row` / `set_table_cell_value` → modify
4. `get_table_data` → verify result

### Add header and footer
1. `get_sections` → see how many sections exist
2. `set_header_footer` with sectionIndex=0, type="header" → set header HTML
3. `set_header_footer` with sectionIndex=0, type="footer" → set footer HTML
4. `get_headers_footers` → verify

### Highlight specific text throughout document
1. `format_found_text` with searchText + desired formatting (bold, color, highlight)
2. `get_document_content` → verify

### Add footnotes/endnotes
1. User selects text in document
2. `insert_footnote` or `insert_endnote` with reference text
3. `get_footnotes_endnotes` → verify

## Always-On Defaults

- Always discover document structure before any modification.
- Always read content before modifying it.
- Prefer `insert_content_at_selection` for rich formatted content.
- Use `insert_paragraph` for simple text additions.
- Always verify changes after mutations.
- Always finish with a clear summary of actions taken.

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue with independent remaining steps.
