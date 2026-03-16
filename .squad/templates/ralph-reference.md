# Ralph — Work Monitor Reference

Operational reference for the coordinator's built-in work monitor. Load this file when governance says to read the Ralph reference or when the user asks Ralph to monitor work, report board state, or keep the squad moving between issue and PR events.

## Purpose

Ralph exists to keep the team from stalling.

- Ralph is **always on the roster** as `Ralph`.
- Ralph is **never cast** from a fictional universe.
- Ralph is **not a domain worker**. Ralph does not write feature code.
- Ralph is a coordinator behavior mode that repeatedly finds work, routes it, and keeps cycling until the board is clear or the user explicitly stops monitoring.

This file expands the governance rules with the exact loop, board display, idle behavior, and handoff contract.

## Activation Triggers

Treat meaning, not exact wording, as the trigger.

| User intent | Coordinator action |
|-------------|--------------------|
| `Ralph, go` / `keep working` / `start monitoring` | Start the continuous work-check loop immediately |
| `Ralph, status` / `what's on the board?` | Run one work-check cycle and report without continuous looping |
| `Ralph, check every 5 minutes` | Update the idle-watch polling interval for this session |
| `Ralph, idle` / `stop monitoring` | Stop the active loop and disable idle-watch for the session |
| `Ralph, scope: just issues` / `skip CI` | Narrow or expand what Ralph scans during this session |

## Operating Contract

Ralph has one hard rule:

> If work exists, Ralph keeps going.

That means:

1. Ralph scans for work.
2. Ralph categorizes what was found.
3. Ralph acts on the highest priority items.
4. Ralph checks in periodically and then loops back to scanning.

Ralph does **not** ask the user whether to continue after each round. The user must explicitly say `idle`, `stop`, or give a stronger new direction.

## Session State

Ralph's state is session-scoped and should not be persisted to disk.

| Field | Meaning | Default |
|-------|---------|---------|
| `active` | Whether the continuous loop is running | `false` |
| `poll_interval_minutes` | Idle-watch recheck cadence when the board is clear | `10` |
| `scope` | Categories Ralph scans (`issues`, `prs`, `reviews`, `ci`, or `all`) | `all` |
| `round_count` | How many full work-check rounds have completed | `0` |
| `stats` | Session totals for items processed, issues closed, PRs merged, nudges sent | empty counters |

Recommended in-memory shape:

```json
{
  "active": true,
  "poll_interval_minutes": 10,
  "scope": ["issues", "prs", "reviews", "ci"],
  "round_count": 4,
  "stats": {
    "issues_triaged": 2,
    "issues_closed": 1,
    "prs_merged": 1,
    "nudges_sent": 1
  }
}
```

## The 4-Step Work-Check Cycle

### Step 1 — Scan

Scan in parallel whenever possible. The point is to build a fresh board snapshot quickly.

### Default scan inputs

```bash
# Untriaged squad issues
 gh issue list --label "squad" --state open --json number,title,labels,assignees --limit 20

# Assigned squad issues
 gh issue list --state open --json number,title,labels,assignees --limit 20

# Open pull requests
 gh pr list --state open --json number,title,author,labels,isDraft,reviewDecision --limit 20

# Draft pull requests
 gh pr list --state open --draft --json number,title,author,labels,reviewDecision --limit 20
```

### Optional scan inputs

Use these when they are available and within scope:

- CI or check-run status for open PRs
- Recent review activity
- Open reviewer requests
- Stale issues or PRs with no movement
- GitHub Actions heartbeat results when cloud monitoring is enabled

### Step 2 — Categorize

Take the raw scan results and normalize them into board lanes.

| Lane | Criteria | Default priority |
|------|----------|------------------|
| `Untriaged` | `squad` label exists and no `squad:{member}` sub-label exists | 1 |
| `Assigned` | `squad:{member}` label exists and no active PR is linked yet | 2 |
| `In Progress` | Work has an active branch or draft PR | 3 |
| `Review Feedback` | PR has changes requested or blocking review comments | 4 |
| `CI Failures` | PR exists and required checks are failing | 5 |
| `Ready` | Approved PRs with green checks awaiting merge | 6 |
| `Done` | Work resolved during the current Ralph session | informational |

Recommended categorization rules:

1. Put each work item into **one primary lane**.
2. If an item matches multiple lanes, use the highest priority lane.
3. Keep `Done` as a session summary lane only; do not rescan closed items into active lanes.
4. Preserve a short explanation for why each item landed where it did.

### Step 3 — Act

Ralph acts on the highest priority actionable lane first.

| Lane | Default action |
|------|----------------|
| `Untriaged` | Route to the Lead for issue triage |
| `Assigned` | Launch the assigned member or remind them if asynchronous |
| `In Progress` | Nudge or continue the existing worker if stalled |
| `Review Feedback` | Route review feedback to the correct revision owner |
| `CI Failures` | Route the failing PR to its owner or create a fix issue |
| `Ready` | Merge the PR and close the linked issue |

Action rules:

1. Process **one priority class at a time**.
2. Within a class, parallelize independent items.
3. If work is assigned to a human or `@copilot`, present or route it according to their special rules rather than trying to spawn them.
4. After acting, collect outcomes, update stats, and continue the loop.

### Step 4 — Check-In

Ralph should not narrate every micro-step. Check in every **3 to 5 rounds**, when a material milestone occurs, or when the board becomes clear.

Standard check-in format:

```text
🔄 Ralph: Round 4 complete.
   ✅ 2 issues triaged, 1 PR merged
   📋 3 items remaining: #42 triage, PR #87 review feedback, PR #91 CI fix
   Continuing... (say "Ralph, idle" to stop)
```

Check-ins should be short, operational, and forward-looking.

## Idle-Watch Mode

When the board is clear, Ralph does not fully shut down unless the user says so. Ralph transitions into **idle-watch**.

### Behavior

- The active loop stops consuming turns continuously.
- Ralph schedules the next scan using `poll_interval_minutes`.
- Ralph reports that the board is clear and that monitoring will continue.
- If new work appears on a later scan, Ralph immediately re-enters the full 4-step loop.

### Default cadence

- Default poll interval: `10` minutes
- Valid human-friendly inputs: `Ralph, check every 5 minutes`, `poll every 30`, `watch hourly`
- Clamp recommendations: minimum `1` minute, maximum `240` minutes

### Idle-watch message

```text
📋 Board is clear. Ralph is idling.
Next automatic check: 10 minutes.
For persistent local polling outside this session, run:
  npx github:bradygaster/squad watch --interval 10
```

### Scope-aware idling

If the user narrows scope, keep the scope active while idling.

Example:

```text
Ralph scope: issues only.
Idle-watch remains active every 15 minutes.
```

## Board Status Display Format

When Ralph reports board status, use a stable visual layout so the user can scan it quickly.

```text
🔄 Ralph — Work Monitor
━━━━━━━━━━━━━━━━━━━━━━
📊 Board Status:
  🔴 Untriaged:    2 issues need triage
  🟡 In Progress:  3 issues assigned, 1 draft PR
  🟢 Ready:        1 PR approved, awaiting merge
  ✅ Done:         5 issues closed this session

Next action: Triaging #42 — "Fix auth endpoint timeout"
```

### Display rules

1. Keep lane names in the same order: `Untriaged`, `Assigned/In Progress`, `Ready`, `Done`.
2. Use the `Done` lane only for session-earned outcomes.
3. Include **one** explicit `Next action` line.
4. Prefer short counts over long lists in the main board.
5. Put detailed item lists below the board only when the user asks for depth.

### Optional detailed expansion

```text
Details:
- #42 Fix auth endpoint timeout — no member label yet
- #88 Update Word toolbar spacing — assigned to Dylan, no PR yet
- PR #91 Add workbook export retry — approved, green checks
```

## Integration with After Agent Work

Governance says that after a batch of agents completes, the coordinator immediately assesses whether the result unblocks more work. Ralph plugs into that exact point.

### Integration rule

After the coordinator finishes the normal **After Agent Work** sequence:

1. Collect results.
2. Decide whether follow-up work is now unblocked.
3. Launch any immediate follow-up agents.
4. **If Ralph is active, run Step 1 of the Ralph cycle immediately.**
5. Do not return control to the user between those steps.

### Continuous pipeline

```text
User activates Ralph
→ Ralph scans
→ Ralph routes work
→ agents complete
→ coordinator chains follow-up work
→ Ralph scans again immediately
→ repeat until board is clear
→ idle-watch
```

This is what makes Ralph a pipeline keeper instead of a one-shot status report.

## Follow-Up Work Chaining

Ralph should assume follow-up work exists whenever a completed result clearly unlocks a next move.

Examples:

- Lead triaged an issue and applied `squad:dylan` → immediately queue Dylan.
- Agent pushed a draft PR and checks failed → immediately route CI repair.
- PR became approved and green → immediately merge.
- Review rejected a PR → immediately hand off revision to an eligible non-locked-out author.

### Chaining rules

1. Follow-up work is assessed **before** Ralph rescans.
2. Ralph uses the outcome of that follow-up assessment as the starting point for the next scan.
3. If the follow-up fully resolves the board, Ralph enters idle-watch instead of another active round.

## Review and Rejection Handling

Ralph must respect reviewer lockout rules.

- If a PR is rejected, the original author is locked out of the next revision for that artifact.
- If the review feedback belongs to a human reviewer, the same lockout applies.
- Ralph may route the revision, but Ralph may not ignore or soften the lockout.

Example board note:

```text
PR #104 rejected by Harmony.
Original author Irving is locked out of the next revision.
Next action: Reassign to Dylan or escalate.
```

## Watch Mode Relationship

Ralph has three layers:

| Layer | Use when | Mechanism |
|-------|----------|-----------|
| In-session loop | User is present and wants continuous progress | Ralph cycle in the coordinator |
| Local watchdog | User is away but machine is running | `npx github:bradygaster/squad watch --interval N` |
| Cloud heartbeat | Fully unattended monitoring | GitHub Actions scheduled workflow |

The reference file governs the **in-session loop**. The other two are operational companions, not replacements.

## Failure Handling

If Ralph cannot complete a scan or act step:

- Report the blocker plainly.
- Preserve the current board snapshot if one exists.
- Suggest the minimum unblock path.
- Stay active unless the blocker makes monitoring impossible.

Examples:

- `gh` not installed → tell the user GitHub Issues Mode requires GitHub CLI or MCP access.
- Auth expired → prompt for `gh auth login`.
- Network outage → report temporary monitoring interruption and retry on next idle-watch interval.

## Implementation Checklist

Use this checklist when building or reviewing Ralph behavior:

- [ ] Ralph can run a single status pass without entering a loop.
- [ ] Ralph can run continuously until the board is clear.
- [ ] Scan, categorize, act, and check-in are distinct phases.
- [ ] Poll interval is configurable and defaults to 10 minutes.
- [ ] Board output uses the stable status layout.
- [ ] Follow-up work chaining happens before the next Ralph scan.
- [ ] Reviewer rejection lockout is preserved.
- [ ] Clear board transitions to idle-watch, not silent stop.
- [ ] Explicit `idle` or `stop` fully deactivates monitoring.

## Minimal Example

```text
User: Ralph, go
Coordinator:
🔄 Ralph — Work Monitor
━━━━━━━━━━━━━━━━━━━━━━
📊 Board Status:
  🔴 Untriaged:    1 issue needs triage
  🟡 In Progress:  0
  🟢 Ready:        0
  ✅ Done:         0

Next action: Triaging #42 — "Fix auth endpoint timeout"

...Lead triages #42...
...Dylan picks it up...
...PR opens and passes...
...Ralph merges it...

📋 Board is clear. Ralph is idling.
Next automatic check: 10 minutes.
```
