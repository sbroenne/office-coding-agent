---
name: word-formatting
description: >
  Specialized skill for text formatting, paragraph styling, and visual refinement
  in Word documents. Covers fonts, colors, spacing, and named styles.
version: 1.0.0
license: MIT
hosts: [word]
---

# Formatting Skill

Activate this skill when formatting, styling, or visually refining document content.

## Font & Inline Formatting

Use `apply_style_to_selection` for inline formatting:
- **bold**, *italic*, underline, strikethrough
- Font name, size, color, highlight color

### Color Values
- Use named colors or hex: `"#4472C4"`, `"red"`, `"#333333"`
- Common professional palette:
  - Headings: `"#1F3864"` (dark blue), `"#333333"` (charcoal)
  - Body: `"#404040"` (dark gray)
  - Accent: `"#4472C4"` (blue), `"#70AD47"` (green), `"#ED7D31"` (orange)

## Paragraph Styling

### Named Styles (`apply_paragraph_style`)
Use built-in Word styles for consistent formatting:
- `Heading 1`, `Heading 2`, `Heading 3` — section hierarchy
- `Normal` — body text
- `Title`, `Subtitle` — document title
- `Quote`, `Intense Quote` — callout blocks
- `List Paragraph` — for bulleted/numbered lists

### Paragraph Format (`set_paragraph_format`)
- **Alignment**: `left`, `center`, `right`, `justified`
- **Spacing**: `spaceBefore`, `spaceAfter` (in points)
- **Line spacing**: `lineSpacing` (in points), `lineUnitBefore`, `lineUnitAfter`
- **Indentation**: `firstLineIndent`, `leftIndent`, `rightIndent` (in points)

## Bulk Formatting with `format_found_text`

Search for text patterns and apply formatting to all matches:
1. Search term + bold/italic/color/highlight
2. Great for highlighting key terms, names, or technical terms throughout the document

## Common Patterns

### Professional Document Formatting
1. `get_document_overview` → understand structure
2. Apply `Heading 1`/`Heading 2` styles to section headings
3. Set body text to consistent font and size
4. Adjust spacing between sections
5. Verify with `get_document_content`

### Highlight Key Terms
1. `format_found_text` with search term + bold + color
2. Repeat for each key term
3. Verify result

## Always Verify
After any formatting change, re-read the affected content to confirm it looks correct.
