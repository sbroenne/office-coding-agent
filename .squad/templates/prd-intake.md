# PRD Intake Reference

Reference for turning a product requirements document into routed Squad work. Load this when the user provides a PRD, points to a spec file, pastes requirements text, or says the PRD has changed mid-project.

## Purpose

PRD Mode lets Squad treat a requirements source as the short-term source of truth for decomposition and prioritization.

The core flow is:

```text
Detect PRD source → store PRD reference in team.md → spawn Lead for decomposition → present work table → route approved items → update when PRD changes
```

This document defines trigger detection, how to capture the PRD source, the Lead decomposition handoff, the work-item table format, and how to handle updates without losing project continuity.

## Triggers for PRD Mode

Enter PRD Mode when the user does any of the following:

| Trigger type | Examples |
|--------------|----------|
| explicit PRD language | `here's the PRD`, `work from this spec`, `read the product doc` |
| file path reference | `read the PRD at docs/prd.md`, `use .squad/specs/launch.md` |
| pasted requirements | multi-paragraph requirements dropped directly into chat |
| URL reference | `use this Notion doc`, `work from this GitHub gist`, `read the spec at https://...` |
| change notice | `the PRD changed`, `updated the spec`, `there is a new version` |

If the message clearly provides product requirements but never says `PRD`, still treat it as PRD Mode.

## Detecting the PRD Source

The coordinator should normalize the source into one of four source kinds.

| Source kind | Detection cue | Stored example |
|-------------|---------------|----------------|
| `file` | repo-relative or absolute path | `docs/prd.md` |
| `paste` | inline prose in the user message | `inline prompt content from 2026-03-16` |
| `url` | `http://` or `https://` reference | `https://example.com/spec` |
| `mixed` | multiple sources provided together | `docs/prd.md + pasted acceptance criteria` |

### Detection rules

1. If the user gives a readable file path, classify as `file`.
2. If the user pastes long-form requirements directly, classify as `paste`.
3. If the user provides a URL and wants it used as the source, classify as `url`.
4. If the user combines sources, keep the primary source plus a note about supplements.

## Store the PRD Reference in `team.md`

After a PRD source is accepted, write a short section in `team.md` so future turns know what the team is building from.

### Canonical section

```md
## PRD Source

- **Type:** file
- **Reference:** docs/prd.md
- **Captured:** 2026-03-16
- **Status:** Active
- **Notes:** Supplemented by acceptance criteria pasted in chat
```

### Field guidance

| Field | Meaning |
|-------|---------|
| `Type` | `file`, `paste`, `url`, or `mixed` |
| `Reference` | the path, URL, or inline identifier |
| `Captured` | when the team started using this source |
| `Status` | `Active`, `Superseded`, or `Draft` |
| `Notes` | optional clarification |

## Spawn the Lead for Decomposition

The Lead owns decomposition because PRD intake is cross-cutting and architectural.

### Spawn guidance

- spawn the Lead synchronously
- use a premium-capable model if model tiering is available
- pass the PRD source and any supporting repo context
- ask for decomposed work items, dependencies, and routing recommendations

### Recommended spawn prompt skeleton

```text
You are the Lead for this squad.
Read the PRD source below and decompose it into executable work items.

PRD SOURCE
Type: file
Reference: docs/prd.md
Captured: 2026-03-16

REQUIRED OUTPUT
1. concise product summary
2. assumptions and open questions
3. work items with priority, estimate, dependencies, and recommended agent owner
4. risks and sequencing notes
```

### What the Lead should return

- a short summary of the product intent
- a list of assumptions or ambiguities
- decomposed work items
- dependency ordering
- recommended owners
- any ceremonies that should happen before work begins

## Work Item Presentation Format

Present decomposed work back to the user in a scan-friendly table before routing the entire batch.

### Canonical work item table

```md
| ID | Priority | Estimate | Agent | Work item | Dependencies | Notes |
|----|----------|----------|-------|-----------|--------------|-------|
| W1 | P0 | M | Harmony | define host routing contract | — | architecture gate |
| W2 | P1 | M | Dylan | implement task pane settings UI | W1 | VS Code parity |
| W3 | P1 | S | Mark | add integration coverage for settings flow | W2 | uses real patterns |
```

### Required columns

| Column | Meaning |
|--------|---------|
| `ID` | stable item identifier for follow-up and updates |
| `Priority` | e.g. `P0`, `P1`, `P2` |
| `Estimate` | e.g. `XS`, `S`, `M`, `L` |
| `Agent` | recommended owner |
| `Work item` | crisp execution-oriented statement |
| `Dependencies` | upstream items that must complete first |
| `Notes` | risk, ceremony need, or acceptance hints |

### Presentation rules

1. Sort by priority, then dependency order.
2. Keep the work item phrasing action-oriented.
3. Include only enough detail to route the work; do not turn the table into a second PRD.
4. If a human approval is needed before routing, say so explicitly below the table.

## Approval and Routing

Default behavior after decomposition:

1. show the work-item table
2. ask the user to approve the plan, trim it, or reprioritize it
3. once approved, route items respecting dependencies
4. fan out independent items in parallel
5. keep Scribe informed through the normal squad flow

Example prompt after decomposition:

```text
Here is the proposed work breakdown from the PRD.
Pick one:
- approve and start
- change priorities
- trim scope
```

## Mid-Project PRD Updates

When the user says the PRD changed, do not throw away the current plan blindly. Reconcile it.

### Update flow

1. read the new PRD source
2. compare it to the existing `## PRD Source` and current work item list
3. spawn the Lead again with both the old and new source references
4. ask the Lead for a delta decomposition
5. preserve stable IDs where the work item is substantively the same
6. mark removed or replaced items as superseded rather than silently erasing them

### Recommended `team.md` update on change

```md
## PRD Source

- **Type:** file
- **Reference:** docs/prd-v2.md
- **Captured:** 2026-03-18
- **Status:** Active
- **Notes:** Supersedes docs/prd.md from 2026-03-16
```

### Delta table format

```md
| ID | Change | Old | New | Action |
|----|--------|-----|-----|--------|
| W2 | modified | settings per host | settings per host + per persona | re-scope Dylan task |
| W4 | added | — | export prompts bundle | add new owner |
| W5 | removed | telemetry panel | — | mark superseded |
```

## Handling Multiple Sources

Sometimes the user gives a PRD file and then pastes clarifications in chat.

Rules:

- keep the file or URL as the primary reference when possible
- capture the chat clarification in the `Notes` field
- pass both into the Lead decomposition prompt
- if the chat clarification contradicts the existing PRD, call out the contradiction rather than guessing

## Risks and Boundaries

PRD intake is about planning and routing, not silent execution.

- Do not auto-route a large project decomposition without showing the work table first.
- Do not invent missing acceptance criteria; surface them as open questions.
- Do not let the PRD override standing governance or reviewer rules.
- Do not lose dependency information between decomposition and routing.

## Example End-to-End Flow

```text
1. User: read the PRD at docs/prd.md
2. Coordinator stores PRD reference in team.md
3. Lead decomposes the PRD into work items
4. Coordinator presents the priority/estimate/agent table
5. User approves
6. Coordinator routes W1 and W4 immediately; W2 waits on W1
7. Later the user says the PRD changed
8. Coordinator re-runs decomposition as a delta and updates the plan
```

## Review Checklist

- [ ] PRD source kind was detected correctly
- [ ] `team.md` includes a `## PRD Source` reference section
- [ ] Lead decomposition prompt includes the source and expected output
- [ ] work is presented in a priority / estimate / agent table before broad routing
- [ ] dependency ordering is visible
- [ ] PRD updates use a delta flow instead of replacing the plan blindly
