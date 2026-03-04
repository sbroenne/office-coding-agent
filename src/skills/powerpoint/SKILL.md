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
| Change slide dimensions       | `set_presentation_size`    |
| Group shapes together         | `group_shapes`             |
| Ungroup a grouped shape       | `ungroup_shapes`           |
| Inspect SmartArt on a slide   | `get_smartart_info`        |

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

## Grouping & Ungrouping Shapes

- Use `get_slide_shapes` first to identify shape indices and types.
- **Group**: `group_shapes` takes a list of shape indices and combines them into a single group. Requires PowerPoint 1.8 API (desktop/web only).
- **Ungroup**: `ungroup_shapes` on a shape with type `Group` releases its children back to the slide.
- After grouping/ungrouping, call `get_slide_shapes` again to see the updated shape list.

## SmartArt

SmartArt graphics cannot be **created** via the Office.js API, but they can be **inspected** and **replaced**:

1. **Inspect** — Use `get_smartart_info` to list SmartArt shapes on a slide (position, size, index).
2. **Read** — SmartArt shapes appear as type `SmartArt` or `Diagram` in `get_slide_shapes`.
3. **Replace** — Delete the SmartArt shape with `delete_shape`, then recreate it visually with `add_slide_from_code` using PptxGenJS shapes, arrows, and text boxes.

### Common SmartArt Patterns (with `add_slide_from_code`)

**Process / Chevron Flow (horizontal)**
```js
const steps = ['Step 1', 'Step 2', 'Step 3'];
const n = steps.length;
const bw = (W - 1) / n - 0.15; // box width with gap
steps.forEach((label, i) => {
  const x = 0.5 + i * (bw + 0.15);
  slide.addShape('rightArrow', { x, y: 2.5, w: bw, h: 1.2, fill: { color: '4472C4' } });
  slide.addText(label, { x, y: 2.5, w: bw, h: 1.2, align: 'center', valign: 'middle', color: 'FFFFFF', fontSize: 14, bold: true, shrinkText: true });
});
```

**Cycle (circular process)**
```js
const items = ['Plan', 'Do', 'Check', 'Act'];
const cx = W / 2, cy = H / 2, r = 1.8;
items.forEach((label, i) => {
  const angle = (i / items.length) * 2 * Math.PI - Math.PI / 2;
  const x = cx + r * Math.cos(angle) - 0.6;
  const y = cy + r * Math.sin(angle) - 0.4;
  slide.addShape('ellipse', { x, y, w: 1.2, h: 0.8, fill: { color: '4472C4' } });
  slide.addText(label, { x, y, w: 1.2, h: 0.8, align: 'center', valign: 'middle', color: 'FFFFFF', fontSize: 13, bold: true, shrinkText: true });
});
```

**Hierarchy / Org Chart**
```js
// Top node
slide.addShape('rect', { x: W/2 - 1, y: 0.5, w: 2, h: 0.6, fill: { color: '4472C4' }, line: { color: 'FFFFFF' } });
slide.addText('CEO', { x: W/2 - 1, y: 0.5, w: 2, h: 0.6, align: 'center', valign: 'middle', color: 'FFFFFF', fontSize: 14, bold: true, shrinkText: true });
// Child nodes + connector lines below
```

**Matrix (4-quadrant)**
```js
const labels = [['Urgent\n+ Important', 'Not Urgent\n+ Important'], ['Urgent\n+ Not Important', 'Not Urgent\n+ Not Important']];
const colors = ['C00000', 'ED7D31', 'FFC000', '70AD47'];
labels.flat().forEach((label, i) => {
  const x = 0.5 + (i % 2) * ((W - 1) / 2 + 0.1);
  const y = 1.5 + Math.floor(i / 2) * ((H - 2) / 2 + 0.1);
  slide.addShape('rect', { x, y, w: (W - 1.2) / 2, h: (H - 2.2) / 2, fill: { color: colors[i] } });
  slide.addText(label, { x, y, w: (W - 1.2) / 2, h: (H - 2.2) / 2, align: 'center', valign: 'middle', color: 'FFFFFF', fontSize: 14, bold: true, shrinkText: true });
});
```

**Pyramid (stacked)**
```js
const tiers = ['Strategy', 'Tactics', 'Operations'];
const totalH = H - 2.5;
tiers.forEach((label, i) => {
  const tierH = totalH / tiers.length;
  const tierW = (W - 1) * (1 - i * 0.25);
  const x = (W - tierW) / 2;
  const y = 0.8 + i * tierH;
  slide.addShape('isoscelesTri', { x, y, w: tierW, h: tierH - 0.05, fill: { color: ['4472C4', '70AD47', 'FFC000'][i] } });
  slide.addText(label, { x, y: y + tierH * 0.25, w: tierW, h: tierH * 0.5, align: 'center', valign: 'middle', color: 'FFFFFF', fontSize: 14, bold: true, shrinkText: true });
});
```

