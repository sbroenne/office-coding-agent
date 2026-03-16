---
name: "adaptive-cards"
description: "How to use the adaptive-cards MCP server for card generation, validation, optimization, and Office-focused card workflows"
domain: "mcp-integration"
confidence: "high"
source: "manual"
tools:
  - name: "generate_card"
    description: "Generate Adaptive Card JSON from a plain-language description"
    when: "Use when you know the user intent but need a first card draft quickly"
  - name: "validate_card"
    description: "Validate Adaptive Card JSON against the schema"
    when: "Use before shipping, saving, or embedding any card payload"
  - name: "data_to_card"
    description: "Turn structured data into an Adaptive Card"
    when: "Use for Excel tables, summary rows, KPI snapshots, or list visualizations"
  - name: "optimize_card"
    description: "Tune a card for a target host"
    when: "Use after authoring when the destination is Outlook, Teams, WebChat, Windows, Viva Connections, or Webex"
  - name: "template_card"
    description: "Apply an Adaptive Card template to data"
    when: "Use for repeatable layouts such as status reports, approvals, or notifications"
  - name: "transform_card"
    description: "Transform a card between host-oriented formats"
    when: "Use when reusing a card concept across different destinations with different rendering expectations"
  - name: "suggest_layout"
    description: "Suggest a layout for content before card generation"
    when: "Use when the content is known but the best card structure is not obvious"
---

## Context

Use this skill when an agent needs to create, inspect, or adapt Adaptive Card payloads through the `adaptive-cards-mcp` MCP server. It is especially useful in this repo when work spans Office hosts and collaboration surfaces: Outlook scenarios that need actionable-message style payloads, Excel scenarios that turn structured data into visual summaries, and cross-host workflows that need a validated card before it is shared with Teams or other supported endpoints.

## Patterns

### Start from intent, then validate

- Use `suggest_layout` first when content is ambiguous or dense.
- Use `generate_card` to create the first draft from the user goal, not from raw JSON.
- Always run `validate_card` before treating the output as complete.
- If the card will be sent to a specific surface, follow with `optimize_card` for that host.

### Prefer `data_to_card` for Excel-shaped data

- Reach for `data_to_card` when the source is tabular: worksheet ranges, table rows, KPI summaries, or exported JSON.
- Keep cards concise. Favor top metrics, short labels, and obvious actions over dumping full worksheets into the card.
- Summarize data first when the source range is wide or noisy; Adaptive Cards work best as a compact view, not a spreadsheet clone.

### Use templates for repeatable Office workflows

- Use `template_card` for recurring patterns like approval cards, incident updates, meeting summaries, and release notifications.
- Keep the template stable and swap the bound data rather than regenerating a whole layout every time.
- Pair templates with `validate_card` so shared layouts do not drift into invalid schema.

### Optimize by destination host

- Outlook benefits from compact cards with clear primary actions, restrained nesting, and conservative schema features.
- Teams and other collaboration surfaces can often tolerate richer layouts, but the safest path is still to validate and optimize for the exact host.
- When reusing a concept across hosts, use `transform_card` rather than hand-editing multiple divergent payloads.

### Office-specific guidance

- **Outlook:** Use Adaptive Cards for actionable-message style concepts, mail workflow payloads, approvals, and concise status updates. Keep actions explicit and text brief.
- **Excel:** Use Adaptive Cards to turn worksheet or table output into digestible summaries for downstream sharing in Outlook or Teams. Treat the card as a visualization layer on top of extracted data.
- **PowerPoint / Word:** These hosts do not render Adaptive Cards directly in the add-in UI, but card generation is still useful when users want to repurpose slide or document summaries into collaboration artifacts.

## Examples

### Example usage patterns

- "Turn this expense approval summary into an Outlook-friendly Adaptive Card, then validate it."
- "Convert the selected Excel sales data into a KPI card with top-line totals and a short trend section."
- "Generate a Teams card for a release update, then optimize the same content for Outlook."
- "Apply our approval template to this JSON payload and validate the result before saving it."

### Practical workflow

1. Gather or summarize the content.
2. Use `suggest_layout` if the structure is unclear.
3. Generate with `generate_card` or `data_to_card`.
4. Apply `template_card` if a reusable layout exists.
5. Run `optimize_card` or `transform_card` for the target host.
6. Finish with `validate_card`.

## Anti-Patterns

- **Skipping validation** — Never assume generated JSON is safe to use without `validate_card`.
- **Overloading the card with spreadsheet detail** — Cards should summarize and guide action, not reproduce every row and column.
- **Ignoring host differences** — A card that looks fine for one surface may be a poor fit for Outlook or another constrained host.
- **Hand-forking multiple host versions too early** — Prefer one source card plus `optimize_card` or `transform_card` to reduce drift.
