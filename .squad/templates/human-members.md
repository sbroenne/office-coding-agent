# Human Team Members Reference

Reference for adding real people to a Squad roster and routing work through them without pretending they are spawnable agents. Load this when the user adds or removes human collaborators, asks work to be routed to a human, or when a human acts as a reviewer.

## Purpose

Humans can sit on the roster beside AI team members, but they behave differently:

- they are identified by real names
- they are not cast from the universe system
- they do not receive `charter.md` or `history.md` files
- the coordinator cannot spawn them through the `task` tool
- work routed to a human must be presented and then held until the human responds through the user

This file defines the roster format, routing behavior, reminders, reviewer semantics, and differences from AI agents.

## Human Member Model

A human member is represented in `team.md` only.

### Canonical roster row

```md
| Brady | Product Manager | — | 👤 Human |
```

Recommended status variants:

| Status text | Use when |
|-------------|----------|
| `👤 Human` | normal active human member |
| `👤 Waiting` | work is currently blocked on their reply |
| `👤 Reviewer` | the human is acting as a designated reviewer |
| `👤 Offline` | known to be away or temporarily unavailable |

## Adding Human Members

Typical user requests:

- `Add Brady as PM`
- `Bridget is joining as designer`
- `Put Stefan on the roster as reviewer`

### Add flow

1. Confirm the human's name and role.
2. Add a row under `## Members` in `team.md`.
3. Update `routing.md` if the role changes routing behavior.
4. Do **not** create `.squad/agents/{name}/charter.md`.
5. Do **not** create `.squad/agents/{name}/history.md`.

Example:

```md
| Stefan | Product Owner | — | 👤 Human |
```

### Removing human members

To remove a human:

1. delete or archive the roster row in `team.md`
2. remove any routing entries that point to them
3. clear any waiting-state reminders for that human
4. do not move folders because no human agent folder exists

## Routing Behavior

Humans are **not spawnable**.

When work routes to a human, the coordinator must:

1. present the work clearly in the chat
2. explain why the human is the correct owner or reviewer
3. wait for the user's relay of that human's response
4. continue unrelated work in parallel if it does not depend on the human

### Presentation format

Use human terms rather than agent terms.

```text
👤 Stefan — Product Owner
Need input on: whether export should keep worksheet formulas or values only.
Why it routes here: product decision with downstream UI and testing impact.
Waiting on Stefan's guidance.
```

### Important constraint

Do **not** imply that the human has been messaged automatically. The coordinator is only presenting the work and pausing for the user's relay.

## Waiting State and Stale Reminders

If a human is holding the critical path, track that waiting state in-session.

### Reminder rule

After **more than one turn** passes without the needed human response, issue a stale reminder.

Required wording pattern:

```text
📌 Still waiting on Stefan for export formula policy.
```

### Reminder behavior

- remind once after the first stale turn
- remind again only when the topic becomes relevant again or the user asks for status
- avoid spamming reminders every turn

### Suggested waiting metadata

```json
{
  "name": "Stefan",
  "thing": "export formula policy",
  "requested_at_turn": 12,
  "last_reminded_at_turn": 14
}
```

This is in-memory coordinator state, not a committed file format.

## Reviewer Rejection Lockout for Humans

A human reviewer has the same rejection authority as an AI reviewer.

If a human rejects an artifact:

1. the original author is locked out of the next revision for that artifact
2. the coordinator must route the revision to a different eligible author
3. the rejected author may not self-revise, pair on, or advise the next revision cycle
4. if all eligible authors are locked out, escalate to the user

### Example

```text
Stefan reviewed PR #88 and rejected the approach.
Irving authored the rejected version, so Irving is locked out of the next revision.
Next action: route the revision to Dylan or Mark, depending on the feedback.
```

## Comparison: Humans vs AI Members

| Dimension | Human member | Spawnable AI member | `@copilot` |
|-----------|--------------|---------------------|------------|
| Cast from universe | No | Usually yes | No |
| Spawnable with `task` | No | Yes | No |
| Charter file | No | Yes | No |
| History file | No | Yes | No |
| Interaction style | coordinator presents work and waits | synchronous spawned work | asynchronous issue pickup |
| Typical use | product decisions, approvals, stakeholder review | coding, testing, design, analysis | bounded issue implementation |
| Badge | 👤 | role emoji | 🤖 |

## Human Routing Patterns

### 1. Human as decision maker

Use when the team needs product, stakeholder, or business clarification.

```text
👤 Brady — Product Manager
Need input on: should the task pane remember the last used model per host?
Impact: settings storage, UX copy, and migration logic.
Waiting on Brady before implementation continues.
```

### 2. Human as reviewer

Use when a human is the approval gate.

```text
👤 Stefan — Reviewer
Review requested for PR #102.
Focus: whether the VS Code look-and-feel matches project standards.
```

### 3. Human as collaborator, not blocker

If the human owns only one slice, continue other work in parallel.

```text
Waiting on Stefan for naming guidance.
Meanwhile Dylan can keep polishing the component layout and Mark can expand the tests.
```

## Team File Guidance

Humans should appear under `## Members` beside the AI roster.

Example mixed roster:

```md
## Members

| Name | Role | Charter | Status |
|------|------|---------|--------|
| Harmony | Lead | `.squad/agents/harmony/charter.md` | 🏗️ Active |
| Dylan | Frontend Dev | `.squad/agents/dylan/charter.md` | ⚛️ Active |
| Stefan | Product Owner | — | 👤 Human |
| Scribe | Session Logger | `.squad/agents/scribe/charter.md` | 📋 Active |
| Ralph | Work Monitor | — | 🔄 Monitor |
```

## Routing File Guidance

Humans may appear in `routing.md` when there is a durable routing rule.

Example:

```md
| Product scope decisions | Stefan 👤 | release criteria, naming decisions, go/no-go calls |
```

Keep routing entries crisp. They tell the coordinator *when to present work to the human*, not how to automate a message.

## Multi-Human Teams

Multiple humans can be on the roster at once.

Rules:

1. Track waiting state independently per human.
2. Route only the decisions that actually belong to that person.
3. Do not treat `any human` as sufficient when a specific reviewer or decision owner is named.
4. If two humans disagree, escalate the conflict to the user rather than inventing a resolution.

## Example Interaction

```text
User: Put Stefan on the team as reviewer.
Coordinator: ✅ Stefan added as 👤 Reviewer.

User: Route the final UX review to Stefan.
Coordinator:
👤 Stefan — Reviewer
Need input on: final VS Code visual parity before merge.
Why it routes here: human visual approval gate.
Waiting on Stefan's feedback.
```

Later:

```text
📌 Still waiting on Stefan for final VS Code visual parity review.
```

## Review Checklist

- [ ] Human member has a roster row with 👤 badge
- [ ] No charter or history files were created for the human
- [ ] Routing presents work rather than pretending to spawn the human
- [ ] Stale reminders trigger after more than one turn of waiting
- [ ] Reviewer rejection lockout is enforced when a human rejects work
- [ ] Independent squad work continues while waiting on human input
