---
name: powerpoint
description: Core PowerPoint skill — tool routing, operating loop, and always-on defaults for all PowerPoint tasks.
version: 2.0.0
license: MIT
hosts: [powerpoint]
---

# PowerPoint Core Skill

Use this as the default orchestration skill for all PowerPoint tasks.

## Operating Loop

1. **Locate** — Call `get_selected_slides` to know which slide the user is on right now.
2. **Discover** — Call `get_presentation_overview` to understand slide count and text content.
3. **Read** — Use `get_presentation_content`, `get_slide_shapes`, or `get_slide_image` to inspect the current slide.
4. **Plan** — Before creating or modifying, choose the right approach for the task.
5. **Execute** — Create, modify, or reorganize slides using the appropriate tool.
6. **Verify** — Use `get_slide_image` to visually inspect the result. Assume there are issues — find them.
7. **Refine** — Fix issues found, then re-verify. Repeat until a full pass reveals no new issues.
8. **Summarize** — Finish with a concise plain-language summary of what was done.

## High-Level Tool Guidance

| Task                          | Primary Tool               |
| ----------------------------- | -------------------------- |
| Understand presentation       | `get_presentation_overview`|
| Read slide text               | `get_presentation_content` |
| See slide visually            | `get_slide_image`          |
| Read speaker notes            | `get_slide_notes`          |
| List shapes with details      | `get_slide_shapes`         |
| List available layouts        | `get_slide_layouts`        |
| Get selected slides           | `get_selected_slides`      |
| Get selected shapes           | `get_selected_shapes`      |
| Add a text box                | `set_presentation_content` |
| Create a rich formatted slide | `add_slide_from_code`      |
| Replace an existing slide     | `add_slide_from_code` with `replaceSlideIndex` |
| Add geometric shape           | `add_geometric_shape`      |
| Add a line/connector          | `add_line`                 |
| Edit text in a shape          | `update_slide_shape` or `set_shape_text` |
| Change shape colors/style     | `update_shape_style`       |
| Move or resize a shape        | `move_resize_shape`        |
| Delete a specific shape       | `delete_shape`             |
| Clear all shapes from slide   | `clear_slide`              |
| Delete a slide                | `delete_slide`             |
| Reorder slides                | `move_slide`               |
| Set slide background color    | `set_slide_background`     |
| Apply a layout to a slide     | `apply_slide_layout`       |
| Copy a slide (text only)      | `duplicate_slide`          |
| Set speaker notes             | `set_slide_notes`          |

## Choosing Between `set_presentation_content` and `add_slide_from_code`

- **`set_presentation_content`**: Quick text box addition. No formatting control. Good for simple annotations.
- **`add_slide_from_code`**: Full PptxGenJS power — text with fonts/colors/sizes, bullet lists, tables, shapes, images. Use this for any slide that needs to look professional.

## Common Workflows

### Summarize a presentation
1. `get_presentation_overview` → get all slide text
2. Provide a concise summary to the user

### Add content to existing slide
1. `get_presentation_content` → read current text
2. `get_slide_shapes` → understand existing shape layout
3. `update_slide_shape` or `set_shape_text` → modify existing text, OR
4. `set_presentation_content` → add a new text box
5. Verify with `get_slide_image`

## Always-On Defaults

- **Always call `get_selected_slides` first** to know the user's current slide.
- Always discover the presentation structure before any modification.
- Prefer `add_slide_from_code` over `set_presentation_content` for user-facing content.
- Use 0-based slide indices consistently.
- **Always verify changes with `get_slide_image` after mutations.**
- Always finish with a clear summary of actions taken.

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue with independent remaining steps.
