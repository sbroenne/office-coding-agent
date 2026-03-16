# `@copilot` Team Member Reference

Reference for adding the GitHub Copilot coding agent to a Squad roster and routing issue work to it safely. Load this when the user asks to add Copilot, when the Lead triages GitHub issues, or when team configuration needs the asynchronous coding agent workflow.

## Purpose

`@copilot` is a special squad member:

- always named `@copilot`
- never cast from a fictional universe
- not spawned through the `task` tool
- activated by issue assignment rather than direct synchronous chat routing

This document defines the roster format, capability profile, auto-assign toggle, routing heuristics, and how the Lead should reason about Copilot as part of the team.

## How to Add `@copilot`

### Step 1 — Confirm the user wants the coding agent

Typical prompt:

```text
Want to include @copilot? It can pick up issues autonomously. (yes/no)
```

### Step 2 — Add the roster entry to `team.md`

Use the coding agent row inside the roster or the dedicated coding-agent section, depending on the project's current structure.

Minimal active row:

```md
| @copilot | Coding Agent | `copilot-instructions.md` | 🤖 Active |
```

### Step 3 — Add the capability profile

The capability profile is required because the Lead uses it during triage.

Recommended section:

```md
## @copilot Capability Profile

| Category | Fit | Notes |
|----------|-----|-------|
| Bug fixes with repro steps | 🟢 | Well-defined, bounded scope |
| Refactoring with tests | 🟡 | Needs squad review |
| Architecture decisions | 🔴 | Requires team judgment |
```

### Step 4 — Add the auto-assign toggle

Store the toggle as an HTML comment in `team.md`:

```html
<!-- copilot-auto-assign: true -->
```

or

```html
<!-- copilot-auto-assign: false -->
```

### Step 5 — Ensure `copilot-instructions.md` exists

`@copilot` does not get a charter. It uses `copilot-instructions.md` plus issue context and the capability profile in `team.md`.

## Comparison: Spawnable Agents vs `@copilot`

| Dimension | Spawnable squad member | `@copilot` issue-assigned member |
|-----------|------------------------|----------------------------------|
| Invocation | Spawn with `task` | Assign issue / label asynchronously |
| Name | Cast or fixed roster name | Always `@copilot` |
| Charter | `charter.md` | `copilot-instructions.md` |
| Work mode | Interactive, session-bound | Autonomous, GitHub issue-driven |
| Branch pattern | Usually `squad/{issue}-{slug}` | Usually `copilot/{issue}-{slug}` |
| Output timing | Returns in-session | Opens commits / draft PR later |
| Reviewer lockout | Enforced through coordinator routing | Enforced during PR review and reassignment |
| Best use | questions, reviews, coordinated changes | bounded issue work with clear specs |

Rule of thumb: if the work must happen **now in this session**, a spawnable member is usually the better choice. If the work is a bounded GitHub issue and can proceed asynchronously, `@copilot` is a candidate.

## Capability Profile Format

The capability profile uses a three-level fit marker.

| Marker | Meaning | Routing rule |
|--------|---------|--------------|
| 🟢 | Good fit | Lead may route directly to `squad:copilot` |
| 🟡 | Needs review | Lead may route to `squad:copilot`, but PR review is mandatory |
| 🔴 | Not suitable | Keep with the human-led squad, not `@copilot` |

### Recommended profile table shape

```md
## @copilot Capability Profile

| Category | Fit | Notes |
|----------|-----|-------|
| Bug fixes with repro steps | 🟢 | Low ambiguity |
| Test additions using existing patterns | 🟢 | Mechanical change |
| Small feature with clear acceptance criteria | 🟡 | Review before merge |
| Refactoring with broad behavioral impact | 🟡 | Needs regression review |
| Architecture / design decisions | 🔴 | Too much judgment |
| Security-sensitive changes | 🔴 | Human ownership required |
```

### Profile authoring rules

1. Keep categories concrete and reviewable.
2. Put the likely high-volume issue classes first.
3. Use the notes column to explain risk, not to repeat the fit color.
4. Keep the profile in `team.md` so triage uses the current project-specific risk tolerance.

## Auto-Assign Behavior

The auto-assign flag controls whether issues labeled `squad:copilot` also get assigned to the GitHub Copilot coding agent automatically.

### Toggle format

```html
<!-- copilot-auto-assign: true -->
```

### Meaning

| Toggle | Behavior |
|--------|----------|
| `true` | When the Lead routes an issue to `squad:copilot`, the automation may assign `@copilot` and let it start |
| `false` | The issue may still be labeled for Copilot, but no automatic assignment occurs until a human chooses it |

### Guidance

- Use `true` when the repo has a clean issue pipeline and the capability profile is already tuned.
- Use `false` when the team is trialing Copilot or wants human gatekeeping before every pickup.

## Routing Details

### Labels

| Label | Meaning |
|-------|---------|
| `squad` | Lead triage inbox |
| `squad:{member}` | Assigned to a named squad member |
| `squad:copilot` | Candidate for autonomous Copilot pickup |

### Lead triage flow for `@copilot`

When the Lead sees an open issue in the `squad` inbox:

1. Read the issue carefully.
2. Compare it against the capability profile.
3. Decide whether it is 🟢, 🟡, or 🔴.
4. If 🟢, route to `squad:copilot` if asynchronous work is acceptable.
5. If 🟡, route to `squad:copilot` only when mandatory review is acceptable and clearly called out.
6. If 🔴, assign to a squad member instead.

### Lead triage guidance

Ask these questions in order:

1. **Is the problem well-defined?** Repro steps, acceptance criteria, or a small diff target favor Copilot.
2. **Is the change bounded?** Single subsystem or low-blast-radius work favors Copilot.
3. **Does it follow established patterns?** Repetitive or patterned work favors Copilot.
4. **Does it require architectural judgment, security review, or product negotiation?** If yes, it is likely 🔴.
5. **Can the team afford asynchronous turnaround?** If the user needs an answer immediately, use a spawnable member instead.

## Good-Fit / Needs-Review / Not-Suitable Examples

### 🟢 Good fit

- fix a failing test with a clear stack trace
- add missing docs for an existing workflow
- update a dependency and regenerate lockfiles
- add a small integration test following a nearby example

### 🟡 Needs review

- add a host-specific tool following a clear existing pattern
- refactor a service that already has strong tests
- implement a small feature where acceptance criteria are explicit

### 🔴 Not suitable

- redesign the VS Code-like task pane experience
- define a new multi-host architecture rule
- change auth, permissions, secret handling, or trust boundaries
- create or modify squad governance, charters, or skill authoring rules

## Issue Pickup and PR Behavior

`@copilot` works through GitHub issues.

Typical flow:

1. Lead applies `squad:copilot`.
2. Automation or human assignment gives the issue to `@copilot` if auto-assign is enabled.
3. `@copilot` creates a `copilot/{issue-number}-{slug}` branch.
4. `@copilot` opens a draft PR.
5. A squad reviewer evaluates the PR.
6. If approved and green, merge using the repository's merge policy.

### PR expectations

- reference the issue (`Closes #123`)
- mention that the task was taken under the squad capability profile
- if the issue was 🟡, explicitly request squad review in the PR body

Suggested PR note for 🟡 work:

```text
⚠️ This task was routed to @copilot under the "needs review" profile. Please have a squad reviewer approve before merge.
```

## Roster Example

```md
## Members

| Name | Role | Charter | Status |
|------|------|---------|--------|
| Harmony | Lead | `.squad/agents/harmony/charter.md` | 🏗️ Active |
| Dylan | Frontend Dev | `.squad/agents/dylan/charter.md` | ⚛️ Active |
| Mark | Tester | `.squad/agents/mark/charter.md` | 🧪 Active |
| Scribe | Session Logger | `.squad/agents/scribe/charter.md` | 📋 Active |
| Ralph | Work Monitor | — | 🔄 Monitor |
| @copilot | Coding Agent | `copilot-instructions.md` | 🤖 Active |

<!-- copilot-auto-assign: true -->
```

## Triage Comment Example

When the Lead routes an issue to Copilot, a short comment helps make the decision visible:

```text
Triage: routing to @copilot.
Reason: clear repro, bounded scope, matches 🟢 capability profile.
Expected follow-up: squad review after draft PR opens.
```

For a rejected Copilot route:

```text
Triage: not routing to @copilot.
Reason: this issue requires architecture and UX judgment, which is 🔴 in the capability profile.
Assigned to Harmony instead.
```

## Interaction with the Rest of the Team

- `@copilot` does not serialize the squad. Other independent work continues.
- Humans and spawnable AI members can still review or extend Copilot-authored PRs.
- If a Copilot PR is rejected, reviewer lockout rules still apply to the next revision owner.
- The Lead owns final routing judgment even when the issue looks like a fit.

## Review Checklist

- [ ] `@copilot` was added to `team.md` with a visible 🤖 role marker
- [ ] capability profile exists and uses 🟢/🟡/🔴
- [ ] `<!-- copilot-auto-assign: true/false -->` exists
- [ ] `copilot-instructions.md` is the referenced instruction file
- [ ] triage guidance is consistent with the capability profile
- [ ] issue routing stays asynchronous rather than using the normal spawn path
