# Ceremony Reference

Reference for defining, triggering, and running squad ceremonies. Load this when the coordinator detects a ceremony request, before/after auto-trigger checks need the exact rules, or when a repo is authoring `.squad/ceremonies.md`.

## Purpose

Ceremonies are structured meetings performed by agents before or after work. They are not free-form chats. Each ceremony should have a clearly defined trigger, facilitator, participant set, and output format.

This file defines the config shape used in `.squad/ceremonies.md`, when ceremonies auto-run vs wait for a manual request, the facilitator spawn template, cooldown rules, and the result presentation format.

## Ceremony Definition Format

`.squad/ceremonies.md` is a human-readable config file. The preferred format is one heading per ceremony followed by a field table and agenda.

### Canonical shape

```md
## Design Review

| Field | Value |
|-------|-------|
| **Type** | design-review |
| **Trigger** | auto |
| **When** | before |
| **Condition** | multi-agent task involving 2+ agents modifying shared systems |
| **Participants** | all-relevant |
| **Facilitator** | lead |
| **Cooldown** | 1 step |
| **Enabled** | ✅ yes |

**Agenda:**
1. Review the task and requirements
2. Agree on interfaces and contracts
3. Identify risks and edge cases
4. Assign action items
```

### Required fields

| Field | Meaning |
|-------|---------|
| `Type` | machine-readable ceremony type or slug |
| `Trigger` | `auto` or `manual` |
| `When` | `before`, `after`, or `manual-only` |
| `Condition` | natural-language matching rule for auto checks |
| `Participants` | who should attend |
| `Facilitator` | who runs the ceremony |
| `Cooldown` | how long auto-trigger rechecks should pause after the ceremony |
| `Enabled` | whether this ceremony is active |

### Optional fields

- `Time budget`
- `Artifacts`
- `Outputs`
- `Notes`

## Config Interpretation Rules

1. A ceremony section title is the human-facing ceremony name.
2. `Type` should remain stable even if the title is lightly edited.
3. `Trigger: auto` means the coordinator evaluates the ceremony during normal routing.
4. `Trigger: manual` means it runs only when the user explicitly asks.
5. `When: before` means check it before spawning the work batch.
6. `When: after` means check it after the work batch completes.
7. `manual-only` is a valid value when a ceremony should never be part of auto routing.

## Auto-Triggered vs Manual Ceremonies

### Auto-triggered ceremonies

Auto-triggered ceremonies are evaluated by the coordinator whenever the relevant stage is reached.

Typical examples:

- design review before a multi-agent architecture change
- retrospective after a failed test run or reviewer rejection
- incident huddle after a production regression

### Manual ceremonies

Manual ceremonies run only when the user asks.

Typical examples:

- `run a retro`
- `start a design review`
- `hold a planning session`

### Matching rule for auto ceremonies

The `Condition` field is semantic, not a strict parser grammar. The coordinator should map the current task to the condition in plain language.

Examples:

| Condition | Likely match |
|-----------|--------------|
| `multi-agent task involving 2+ agents modifying shared systems` | full-stack feature touching UI and backend |
| `build failure, test failure, or reviewer rejection` | a post-failure retrospective |
| `new public API or schema change` | contract review before implementation |

## Facilitator Selection

The `Facilitator` field names the role or member responsible for running the ceremony.

Common values:

- `lead`
- `tester`
- `all-relevant`
- explicit member name such as `Harmony`

Guidance:

- choose the Lead for decisions, interfaces, and prioritization
- choose the Tester for retrospectives focused on quality or failure analysis
- choose an explicit named facilitator when the ceremony belongs to one domain owner

## Facilitator Spawn Template

The facilitator is spawned synchronously. The facilitator may then fan out to participants if needed.

### Canonical template

```text
You are facilitating the "{Ceremony Name}" ceremony.

CEREMONY CONTEXT
Type: {type}
Trigger: {auto|manual}
When: {before|after|manual-only}
Condition matched: {why this ceremony is running now}
Participants: {participant set}
Facilitator: {facilitator}
Input work: {task, issue, PR, or batch summary}

YOUR JOB
1. summarize the problem or work batch
2. collect the relevant viewpoints from the listed participants
3. produce decisions, action items, and owners
4. keep the result concise and operational

OUTPUT FORMAT
- Summary
- Decisions
- Action items
- Risks / follow-ups
```

### Facilitator behavior rules

1. The facilitator owns structure and pacing, not the domain truth alone.
2. The facilitator may spawn participant agents as sub-tasks if needed.
3. The facilitator should resolve obvious alignment questions during the ceremony rather than punting them all back to the user.
4. The facilitator should produce a compact result that can be included in later spawn prompts.

## Execution Rules

### Before-work ceremonies

1. detect the matching `before` ceremony
2. run the ceremony before spawning the work batch
3. include the resulting decisions or contracts in the work prompts
4. start the work batch only after the ceremony completes

### After-work ceremonies

1. detect the matching `after` ceremony once the work batch or review result is known
2. run the ceremony against the completed batch or failure context
3. present the result and any new action items
4. apply cooldown before checking again automatically

### Manual ceremonies

1. when the user explicitly requests a ceremony, run it even if no auto condition matches
2. manual execution does not require a matching `before`/`after` stage
3. still apply cooldown if the ceremony would otherwise immediately re-trigger automatically

## Cooldown Rules

Ceremony cooldown prevents the same auto ceremony from firing repeatedly on adjacent steps.

### Standard cooldown

Default cooldown is **1 immediately following step** unless the ceremony config says otherwise.

### Why cooldown exists

Without cooldown:

- a `before` design review might re-run on every subtask spawn
- an `after` retrospective might fire repeatedly on the same failure loop

### Cooldown behavior

- cooldown applies only to auto-trigger checks
- manual requests bypass cooldown
- once the cooldown window passes, the ceremony becomes eligible again if the condition still matches

### Recommended representation

```json
{
  "ceremony_type": "design-review",
  "cooldown_steps_remaining": 1
}
```

This is coordinator session state, not a committed file requirement.

## Ceremony Result Presentation Format

After a ceremony completes, present a short result summary.

### Required summary line

```text
📋 Design Review completed — facilitated by Harmony. Decisions: 3 | Action items: 4.
```

### Recommended detailed expansion

```md
**Decisions**
- Keep host routing in a shared service
- Add integration coverage before UI polish
- Require Lead review on cross-host contract changes

**Action items**
- Dylan: implement task pane update flow
- Irving: refactor host tool selection
- Mark: add integration coverage for routing
```

### Result rules

1. Count decisions and action items explicitly.
2. Name the facilitator.
3. Keep the summary concise enough to include inline with normal coordinator narration.
4. Feed action items back into normal routing logic after the ceremony completes.

## Example Ceremony Definitions

### Auto `before` design review

```md
## Design Review

| Field | Value |
|-------|-------|
| **Type** | design-review |
| **Trigger** | auto |
| **When** | before |
| **Condition** | multi-agent task involving 2+ agents modifying shared systems |
| **Participants** | all-relevant |
| **Facilitator** | lead |
| **Cooldown** | 1 step |
| **Enabled** | ✅ yes |
```

### Auto `after` retrospective

```md
## Retrospective

| Field | Value |
|-------|-------|
| **Type** | retrospective |
| **Trigger** | auto |
| **When** | after |
| **Condition** | build failure, test failure, or reviewer rejection |
| **Participants** | all-involved |
| **Facilitator** | lead |
| **Cooldown** | 1 step |
| **Enabled** | ✅ yes |
```

### Manual planning session

```md
## Sprint Planning

| Field | Value |
|-------|-------|
| **Type** | planning |
| **Trigger** | manual |
| **When** | manual-only |
| **Condition** | user explicitly asks for planning |
| **Participants** | lead, product, relevant implementers |
| **Facilitator** | lead |
| **Cooldown** | 0 |
| **Enabled** | ✅ yes |
```

## Scribe Interaction

Ceremonies often produce durable decisions. After a substantive ceremony:

- include Scribe in the background if the normal workflow already calls for it
- capture final decisions in the usual decisions pipeline
- do not treat ceremony notes as authoritative until they enter the project's normal decision flow

## Review Checklist

- [ ] `.squad/ceremonies.md` uses the heading + field-table format
- [ ] each ceremony has name, type, trigger, participants, facilitator, and condition
- [ ] auto vs manual behavior is explicit
- [ ] facilitator prompt follows the reference template
- [ ] cooldown prevents immediate repeated auto triggers
- [ ] ceremony results use the `📋 ... completed` summary format
