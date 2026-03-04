---
name: powerpoint-speaker-notes
description: Auto-generate speaker notes for every slide — talking points, timing cues, and presentation guidance.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# Speaker Notes Skill

Generate high-quality speaker notes for every slide using `set_slide_notes`.

## When to Generate Notes

**ALWAYS** generate speaker notes after:

- Creating a new slide
- Significantly modifying slide content (new data, restructured layout, changed message)
- Building a multi-slide deck (add notes to each slide before moving to the next)

Call `set_slide_notes` with `slideIndex` (0-based) and `notes` (plain text string) immediately after each slide creation or major edit.

## Note Structure

Every note should follow this four-part structure:

1. **Opening hook** — 1 sentence to engage the audience and set up the slide's topic.
2. **Key talking points** — 3–5 bullets, each 1–2 sentences. These are the core of the note.
3. **Transition** — 1 sentence bridging to the next slide's topic.
4. **Timing suggestion** — e.g., "~2 minutes" so the presenter can pace themselves.

## Writing Style

- **Conversational, not scripted.** Write the way a confident presenter speaks — not how they read.
- **Address the audience with "you."** ("You'll notice that…", "Ask yourself…")
- **Include rhetorical questions.** ("So why does this matter?", "What would happen if…?")
- **Mark pauses.** Use "[PAUSE]" where the presenter should let a point land before continuing.
- **Keep it glanceable.** A presenter looks down for 2 seconds — they need to find their place instantly.

## Content Guidelines

- **Add context not visible on the slide.** If the slide says "Revenue grew 34%", the note should explain *why* and what it means for the audience.
- **Include data sources or citations.** ("Source: Gartner 2024 Magic Quadrant", "Per our Q3 internal audit")
- **Suggest audience interaction.** ("Good moment to ask: 'Has anyone experienced this?'")
- **Anticipate objections.** ("If someone pushes back on cost, note that ROI breaks even in 8 months per the finance model.")

## Anti-Patterns

- ❌ **Don't repeat slide text.** If a bullet says "Reduced churn by 15%", the note should NOT say "We reduced churn by 15%." Instead explain the how or the so-what.
- ❌ **Don't write a full script.** Notes over 150 words become a crutch. Keep each note under 150 words.
- ❌ **Don't include formatting.** Speaker notes are plain text — no markdown, bold, or bullets syntax. Use line breaks and dashes for structure.
- ❌ **Don't skip notes.** Every slide deserves notes, even simple divider slides ("Pause here — take a breath before the next section.").

## Example Notes

### 1. Title / Intro Slide

Slide content: "Accelerating Growth in APAC — Q4 Strategy Review"

```
set_slide_notes call:
  slideIndex: 0
  notes:
```

```
Welcome everyone — thanks for making time today. This session is about one thing: how we unlock the next wave of growth in Asia-Pacific before year-end.

- You'll recall we set a target of 20% YoY revenue growth in APAC at the start of FY24. We're tracking at 14% — solid, but behind plan.
- Today I'll walk you through three levers we're pulling in Q4 to close that gap.
- By the end of this deck you'll know exactly where we're doubling down and where we're pivoting.

[PAUSE] — let that target gap sink in before moving on.

Next up, I'll show you where we stand today with the regional dashboard.

~1 minute
```

### 2. Data / Chart Slide

Slide content: Bar chart showing monthly active users by region (NA, EMEA, APAC) over 6 months.

```
set_slide_notes call:
  slideIndex: 3
  notes:
```

```
The chart tells one clear story: APAC is accelerating while NA has plateaued. So why does this matter?

- APAC grew 42% in 6 months — faster than NA did in its first two years. Source: Product Analytics, Oct 2024 export.
- NA's plateau isn't a problem yet, but if APAC follows the same curve, it overtakes NA by Q2 next year.
- The EMEA dip in August correlates with the EU pricing change — not churn. Retention held at 91%.

Good moment to ask the room: "Which region do you think has the highest expansion revenue per user?" [PAUSE] — the answer is APAC, and I'll show that on the next slide.

If someone questions the APAC spike, note that it excludes the Japan enterprise deal to avoid skewing the trend.

~2 minutes
```

### 3. Key Takeaway / Closing Slide

Slide content: "Three things to remember" with three short phrases.

```
set_slide_notes call:
  slideIndex: 11
  notes:
```

```
This is the slide they'll remember — so slow down here.

- First: APAC is our biggest growth lever. Every dollar invested there returns 3x what NA does right now. Source: FP&A model v4.2.
- Second: product-led growth is working. Free-to-paid conversion hit 8.3% — double the industry benchmark.
- Third: we need to hire 12 more AEs in Singapore and Sydney by January or we leave pipeline on the table.

[PAUSE] — let silence do the work. Then ask: "Which of these three feels most urgent to you?" Take 2-3 responses before wrapping up.

Thank everyone for their time and point them to the shared doc for next steps.

~2 minutes
```

## Reminder

After every `add_slide_from_code`, `set_presentation_content`, or major `update_slide_shape` call, follow up with:

```
set_slide_notes({ slideIndex: <0-based index>, notes: "<plain text notes>" })
```

Never leave a slide without speaker notes.
