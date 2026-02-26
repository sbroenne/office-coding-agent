---
name: powerpoint-deck-builder
description: >
  Specialized skill for creating new slide decks. Covers layout variety,
  verification loops, and multi-slide deck workflows.
version: 1.1.0
license: MIT
hosts: [powerpoint]
---

# Deck Builder Skill

Activate this skill when creating new presentations or adding multiple slides.

## Deck Creation Workflow

1. `get_presentation_overview` → understand current state
2. **Plan layout types** for each slide before creating any (vary layouts!)
3. For each slide: create → verify → fix → verify → next slide
4. Summarize what was created

## Layout Variety

**Do NOT default to title + bullet slides for everything.** Mix layouts:

- Title slides, bullet lists, two-column, three-column cards
- Full-bleed color dividers, stat/number callouts, quote slides, tables
- Match content to layout: comparisons → columns, metrics → stat callout, quotes → centered

**Rule:** Never use the same layout for more than 2 consecutive slides.

## Create → Verify → Fix Loop

**This is the most important part. Run it for EVERY slide.**

```
For each slide {
  1. Create with add_slide_from_code
  2. get_slide_image(region: "full") — overview check
  3. get_slide_image(region: "bottom-left") + get_slide_image(region: "bottom-right")
     → Zoomed 2x detail where text overflow happens
  4. If ANY issue → fix → repeat from step 2
  5. Move to next slide only when it looks right
}
```

### Common fixes:
| Problem | Fix |
|---------|-----|
| Text cut off at bottom | Shorten text or remove a bullet |
| Text too small | Increase fontSize, reduce content to fit |
| Word breaking mid-word | Use shorter synonym ("Medikamentenentwicklung" → "Arzneimittel") |
| Too cramped | Reduce content or use fewer columns (4→3) |
| Too many bullets with intro | Remove least important bullet |

### Key principle:
**If something looks wrong in `get_slide_image`, fix it and look again.** Don't move on until it looks right. Expect 1-3 fix cycles per slide.

## Content Tips

- **Keep text short** — punchy phrases, not full sentences
- **Prefer 3 columns** over 4 — gives more room
- **`shrinkText: true`** on all `addText()` as safety net
- **If it overflows, shorten the text** — that's better than tiny fonts
