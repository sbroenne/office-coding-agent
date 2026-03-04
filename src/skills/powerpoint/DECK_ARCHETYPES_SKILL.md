---
name: powerpoint-deck-archetypes
description: Common presentation templates — pitch deck, quarterly review, project status, training deck with slide-by-slide structure.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# Deck Archetypes

Use these slide-by-slide blueprints when the user asks for a standard presentation type. Each archetype specifies layout types, content density, and visual variety rules.

## Global Rules

### Content Density

| Slide type       | Max words | Notes                        |
| ---------------- | --------- | ---------------------------- |
| Title slide      | 12        | Company name + tagline only  |
| Bullet slide     | 40        | 3-5 bullets, ≤ 8 words each |
| Chart slide      | 20        | Title + axis labels + legend |
| Table slide      | 50        | Short cell values            |
| 2-column layout  | 50        | ~25 words per column         |
| 3-column cards   | 45        | ~15 words per card           |
| Stat callout     | 15        | Big number + one-liner       |

### Visual Variety

- **Never use the same layout for more than 2 consecutive slides.**
- Alternate between text-heavy and visual slides (chart, table, stat callout).
- Every deck should include at least one chart or table slide.
- Use section dividers (full-bleed color slide with large heading) to separate major topics in decks > 8 slides.

### Color Accent Strategy

- Pick **one primary accent** color from the brand palette (default: `4472C4`).
- Use **one secondary accent** for contrast (default: `ED7D31`).
- Reserve **red** (`C00000`) for risk / negative metrics only.
- Reserve **green** (`70AD47`) for positive metrics / success states.
- Backgrounds stay neutral (white or dark theme) — accent color goes on shapes, headings, and chart bars.

---

## Archetype 1 — Pitch Deck (10-12 slides)

Use when: founder pitch, investor meeting, startup overview.

| Slide # | Title                  | Layout           | Content guidance                                                      |
| ------- | ---------------------- | ---------------- | --------------------------------------------------------------------- |
| 1       | Title                  | Title slide      | Company name, tagline (≤ 6 words), optional logo placeholder          |
| 2       | Problem                | Bullets          | Pain-point headline + 3 bullet points describing the problem          |
| 3       | Solution               | 3-column cards   | Hero statement as heading + 3 feature cards (icon placeholder + text) |
| 4       | How It Works           | 3-step process   | Numbered chevron/arrow flow — one short phrase per step               |
| 5       | Market Opportunity     | Chart + text     | TAM / SAM / SOM as nested rectangles or bar chart + one-liner labels  |
| 6       | Business Model         | 2-column layout  | Revenue streams left, pricing tiers or unit economics right           |
| 7       | Traction & Milestones  | Timeline / stats | Horizontal timeline with 4-5 milestones OR stat callout grid         |
| 8       | Team                   | 3-4 column cards | Name, role, one-line bio per person; photo placeholder circles        |
| 9       | Competitive Landscape  | Table            | Comparison matrix — rows = features, columns = competitors + us      |
| 10      | Financial Projections  | Chart            | Grouped bar or line chart — 3-year revenue / growth projection        |
| 11      | The Ask                | 2-column layout  | Left: funding amount (stat callout). Right: use-of-funds breakdown   |
| 12      | Contact / Closing      | Title slide      | "Thank you", contact email, website URL                               |

### Pitch Deck Tips

- Slides 2-3 are the emotional hook — keep them punchy, not wordy.
- Slide 5 (market) must include a chart; do not use bullets for TAM/SAM/SOM.
- Slide 9 (competition) uses a table with checkmarks (✓) and dashes (—) for feature comparison.
- Slide 11 (the ask) should have a single large number for the funding amount.

---

## Archetype 2 — Quarterly Business Review (8-10 slides)

Use when: QBR, executive update, board review, performance readout.

| Slide # | Title                      | Layout           | Content guidance                                                        |
| ------- | -------------------------- | ---------------- | ----------------------------------------------------------------------- |
| 1       | Title                      | Title slide      | "Q[N] Business Review", company name, date                              |
| 2       | Executive Summary          | Bullets          | 3-4 high-level takeaways — what went well, what needs attention         |
| 3       | Revenue / KPI Dashboard    | Chart + stats    | 2-3 stat callouts (revenue, growth %, target %) + trend chart           |
| 4       | Key Achievements           | 3-column cards   | Top 3 wins with short descriptions                                      |
| 5       | Challenges & Risks         | Table            | Risk register table — risk, impact (H/M/L), mitigation                  |
| 6       | Department Update 1        | 2-column layout  | Metrics left, narrative highlights right                                 |
| 7       | Department Update 2        | 2-column layout  | Vary department; swap column order for visual variety                    |
| 8       | Next Quarter Goals         | Bullets          | 3-5 SMART goals with owners                                             |
| 9       | Budget / Financial Summary | Chart            | Actual vs. plan bar chart; highlight variances                           |
| 10      | Q&A / Discussion           | Title slide      | "Questions?" or "Discussion" with contact info                           |

### QBR Tips

- Open with the executive summary so leadership gets the headlines immediately.
- Slide 3 must use a chart — never list KPIs as plain bullets.
- Challenges slide (5) uses red accent for high-impact risks.
- Keep department updates to 2 slides max; merge smaller teams into one slide.

---

## Archetype 3 — Project Status Report (6-8 slides)

Use when: project update, sprint review, program status, steering committee.

| Slide # | Title                  | Layout           | Content guidance                                                       |
| ------- | ---------------------- | ---------------- | ---------------------------------------------------------------------- |
| 1       | Title                  | Title slide      | Project name, status date, overall RAG status (Red/Amber/Green badge)  |
| 2       | Executive Summary      | Stat callout     | 3-4 stat boxes — % complete, days remaining, open risks, budget used   |
| 3       | Timeline / Milestones  | Timeline         | Horizontal timeline or Gantt-style bars — completed vs. upcoming       |
| 4       | Budget Status          | Chart + table    | Budget bar chart (planned vs. actual vs. forecast) + summary table     |
| 5       | Risk Register          | Table            | Columns: risk, probability (H/M/L), impact (H/M/L), mitigation, owner |
| 6       | Key Decisions Needed   | Bullets          | 2-3 decision items with options and recommendation                     |
| 7       | Next Steps             | Bullets          | Action items with owners and target dates                              |
| 8       | Appendix (optional)    | Table or chart   | Detailed data, burn-down chart, or resource allocation                 |

### Project Status Tips

- Slide 1 title slide should include a prominent RAG indicator (colored circle or badge).
- Use green/amber/red shapes for status indicators — not just text colors.
- Timeline (slide 3) should visually distinguish completed milestones (filled) from upcoming (outline).
- Budget chart (slide 4) uses grouped bars: planned (gray), actual (accent), forecast (dashed outline).

---

## Archetype 4 — Training / Workshop Deck (10-15 slides)

Use when: onboarding, training session, workshop, lunch-and-learn, enablement.

| Slide # | Title                     | Layout           | Content guidance                                                       |
| ------- | ------------------------- | ---------------- | ---------------------------------------------------------------------- |
| 1       | Title                     | Title slide      | Training title, presenter name, date                                   |
| 2       | Agenda                    | Bullets          | Numbered list of 4-6 session topics with estimated durations           |
| 3       | Learning Objectives       | 3-column cards   | 3 objectives as cards: "By the end you will be able to…"              |
| 4       | Section Divider 1         | Full-bleed color | Section title in large white text on accent-colored background         |
| 5       | Content — Concept         | 2-column layout  | Concept explanation left, diagram or visual right                      |
| 6       | Content — Example         | Chart or table   | Worked example with real data — chart, table, or code snippet          |
| 7       | Exercise / Activity       | Bullets + callout| Instructions for hands-on activity; time box in a callout shape        |
| 8       | Section Divider 2         | Full-bleed color | Next section title (use secondary accent to vary from divider 1)       |
| 9       | Content — Deep Dive       | 2-column layout  | Additional concept; swap column order from slide 5                     |
| 10      | Content — Best Practices  | 3-column cards   | 3 best-practice cards with do/don't framing                            |
| 11      | Exercise / Discussion     | Bullets + callout| Group exercise or discussion prompt with time box                      |
| 12      | Summary / Key Takeaways   | Bullets          | 3-5 takeaway bullets restating learning objectives                     |
| 13      | Resources & Further Reading| Table           | Resource name, link/location, description columns                      |
| 14      | Q&A                       | Title slide      | "Questions?" with presenter contact info                               |
| 15      | Thank You (optional)      | Title slide      | Closing message, feedback survey link                                  |

### Training Deck Tips

- Use section dividers to break the deck into logical modules (alternate primary / secondary accent).
- Exercise slides should include a visible time-box callout (e.g., "⏱ 10 minutes").
- Content slides should alternate between 2-column and 3-column layouts to maintain visual interest.
- Slide 3 (learning objectives) is critical — frame each objective as a measurable outcome.
- Keep bullet slides to ≤ 5 items; if you need more, split across two slides.

---

## Selecting an Archetype

When the user requests a presentation, match to an archetype:

| User intent keywords                              | Archetype              |
| ------------------------------------------------- | ---------------------- |
| pitch, investor, startup, funding, raise           | Pitch Deck             |
| QBR, quarterly, review, board, executive update    | Quarterly Business Review |
| status, project, update, sprint, steering          | Project Status Report  |
| training, workshop, onboarding, teach, learn       | Training / Workshop    |

If the user's request doesn't clearly match an archetype, ask which template is closest before proceeding. When in doubt, start with the archetype structure and adapt slide count and content to fit the user's specific needs.
