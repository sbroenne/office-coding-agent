---
name: powerpoint-formatting
description: >
  PptxGenJS formatting rules and code patterns for add_slide_from_code.
  Covers text formatting, color conventions, anti-patterns, and correct code examples.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# Formatting Skill

Activate this skill when creating or fixing slides with `add_slide_from_code`. Contains PptxGenJS best practices and anti-patterns.

## PptxGenJS Best Practices

### Text structure
- **Bold all headings and inline labels**: Use `bold: true` for slide titles, section headers, and labels
- **Consistent bullet style**: Use `{ bullet: true }` or `{ bullet: { type: "number" } }` — don't use unicode bullets (•, ‣, etc.)
- **Multi-item content**: Create separate array items for each bullet/paragraph — never concatenate into one string
- **Bold label + description**: ALWAYS combine into ONE string with colon: `"Label: Description"`. NEVER use separate text runs for bold label + normal description — they merge without spacing (renders "LabelDescription"). NEVER use nested text arrays — they render as `[object Object]`.

### Colors and positioning
- **Color values**: Use 6-digit hex without # prefix: `"4472C4"` not `"#4472C4"`
- **Positioning**: All x, y, w, h values are in inches. Check `get_presentation_overview` for actual slide dimensions.
- **Safe margins**: Keep content within 0.5" from all slide edges (e.g. on 16:9: x ≥ 0.5, y ≥ 0.5, x+w ≤ 12.83, y+h ≤ 7.0)
- **Leave 0.3" buffer at bottom** — never fill content to exactly y+h = 7.0"
- **Always add `shrinkText: true`** to every `addText()` call — auto-shrinks text on overflow

### Font sizes
| Element | Font Size |
|---------|-----------|
| Slide title | 28–36pt |
| Subtitle | 18–22pt |
| Body text / bullets | 14–16pt |
| Card/column content | 11–13pt |
| Table cells | 11–13pt |
| Captions / labels | 10–12pt |

## Anti-Patterns — NEVER DO THESE

### ❌ All items in one text element
```js
slide.addText("Step 1: Do the first thing. Step 2: Do the second thing.", { x: 0.5, y: 2, w: 9, h: 4, fontSize: 18 });
```

### ❌ Separate bold label + normal description (renders "Machine LearningSystems that…")
```js
{ text: "Machine Learning", options: { bold: true, bullet: true } },
{ text: "Systems that learn from data", options: { fontSize: 12 } },
```

### ❌ Nested text array (renders as "[object Object],[object Object]")
```js
{ text: [
  { text: "Machine Learning: ", options: { bold: true } },
  { text: "Systems that learn from data" }
], options: { bullet: true } },
```

### ❌ Unicode bullets instead of PptxGenJS bullets
```js
slide.addText("• First point\n• Second point", { x: 0.5, y: 2, w: 9, h: 3 });
```

### ❌ Hash prefix on colors
```js
slide.addText("Title", { color: "#4472C4" });  // WRONG — use "4472C4"
```

## Correct Patterns

### ✅ Flat array with simple strings
```js
slide.addText([
  { text: "Step 1: Do the first thing", options: { bullet: true, fontSize: 16 } },
  { text: "Step 2: Do the second thing", options: { bullet: true, fontSize: 16 } },
], { x: 0.5, y: 2, w: 9, h: 4, shrinkText: true });
```

### ✅ Label with colon in single string (THE ONLY correct way for label+description)
```js
slide.addText([
  { text: "Machine Learning: Systems that learn from data", options: { bullet: true, fontSize: 14 } },
  { text: "Computer Vision: Machines interpreting visual info", options: { bullet: true, fontSize: 14 } },
], { x: 0.5, y: 2, w: 9, h: 4, shrinkText: true });
```

### ✅ Title + subtitle
```js
slide.addText("Title", { x: 0.5, y: 0.5, w: 9, h: 1, fontSize: 32, bold: true, color: "363636" });
slide.addText("Subtitle", { x: 0.5, y: 1.6, w: 9, h: 0.6, fontSize: 18, color: "666666" });
```

### ✅ Table
```js
slide.addTable([["Header 1", "Header 2"], ["Row 1", "Data"]], { x: 0.5, y: 2, w: 9, fontSize: 13 });
```

### ✅ Shape with fill
```js
slide.addShape("rect", { x: 1, y: 1, w: 3, h: 1, fill: { color: "4472C4" } });
```
