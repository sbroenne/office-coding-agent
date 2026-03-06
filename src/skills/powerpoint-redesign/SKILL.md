---
name: powerpoint-redesign
description: >
  Specialized skill for redesigning existing slides and adapting templates.
  Covers shape manipulation, template adaptation pitfalls, and redesign workflows.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# Redesign Skill

Activate this skill when redesigning existing slides, adapting templates, or modifying the visual layout of existing content.

## Redesign Workflow

1. `get_selected_slides` → which slide
2. `get_slide_image` → see current visual design
3. `get_presentation_content` → read the text content
4. `get_slide_shapes` → understand shape layout and positions
5. `add_slide_from_code` with `replaceSlideIndex` → first draft
6. `get_slide_image` → inspect result (assume issues exist — find them)
7. Refine and re-verify until polished

## Template Modification Workflow

1. `get_presentation_overview` → full structure
2. `get_slide_image` on each relevant slide → visual analysis
3. `get_slide_layouts` → available layouts in this deck
4. Plan which slides to keep, modify, delete, or add
5. Make structural changes first (delete/add/reorder slides)
6. Then edit content on each slide
7. Verify each modified slide with `get_slide_image`

## Template Adaptation Pitfalls

### Content count mismatch
When source content has fewer items than the template expects:
- **Remove excess elements entirely** — don't just clear text from shapes
- Use `delete_shape` to remove unneeded shapes
- Use `get_slide_shapes` to identify what to remove
- Verify visually that layout still works after removal

When source content has more items than space allows:
- Split across multiple slides rather than cramming
- Consider truncating or summarizing to fit

### Text length mismatch
- **Shorter replacements**: Usually safe
- **Longer replacements**: May overflow text boxes or wrap unexpectedly
- Always verify with `get_slide_image` after text changes
- Adjust font size or box dimensions with `move_resize_shape` if needed

## Shape Manipulation Guide

### Reading shapes
- `get_slide_shapes` → full list with indices, types, positions, sizes
- Use shape indices (0-based) for targeted modifications

### Modifying shapes
- `update_slide_shape` or `set_shape_text` → change text content
- `update_shape_style` → change fill color, border, font properties
- `move_resize_shape` → reposition or resize (x, y, w, h in inches)
- `delete_shape` → remove a specific shape by index

### When to replace vs. modify
- **Modify** when changing text or minor styling (faster, preserves other shapes)
- **Replace** with `add_slide_from_code` + `replaceSlideIndex` when the layout needs fundamental changes
- Always `get_slide_image` after replacement to verify

## Iterative Verification — MANDATORY

**You MUST call `get_slide_image` after EVERY modification. No exceptions.**

1. Make changes
2. `get_slide_image` → inspect result
3. Check: text cut off? Words breaking? Overlap? Missing content?
4. If ANY issue → fix with `replaceSlideIndex` → go back to step 2
5. Only declare done when a full pass shows zero issues

Expect 2-3 fix cycles per slide. If you verify a slide only once, you probably missed something.
