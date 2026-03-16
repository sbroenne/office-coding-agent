# GitHub Issue Lifecycle Reference

Reference for Squad's GitHub Issues Mode from initial repository connection through branch creation, PR review, and merge. Load this when the user connects a repo, asks the team to work the backlog, or when Ralph handles issue and PR flow.

## Purpose

GitHub Issues Mode lets Squad treat GitHub as an external work queue.

The lifecycle is:

```text
Issue picked up → branch created → work done → PR opened → reviewed → merged
```

This document defines the connection record in `team.md`, the spawn prompt additions for issue work, PR review handling, and the standard merge commands.

## Prerequisites

Before a repo is connected, verify access.

### Preferred access path

1. Use GitHub MCP tools if they are available.
2. Otherwise use the `gh` CLI.

### `gh` checks

```bash
gh --version
gh auth status
```

If `gh` is missing or unauthenticated, stop and tell the user what to install or configure.

## Repository Connection Format

When a repository is connected, store the reference in `team.md` under `## Issue Source`.

### Canonical section

```md
## Issue Source

- **Repository:** owner/repo
- **Connected:** 2026-03-16
- **Filters:** All open issues with `squad` label
- **Auto-assign @copilot:** Enabled
```

### Required fields

| Field | Meaning |
|-------|---------|
| `Repository` | GitHub repository in `owner/repo` format |
| `Connected` | date the team began using this issue source |
| `Filters` | which issues Ralph and the coordinator should treat as the active backlog |
| `Auto-assign @copilot` | whether `squad:copilot` issues are automatically assigned |

### Connection examples

```text
pull issues from microsoft/vscode
connect to sbroenne/office-coding-agent
work on issues from octo-org/design-system
```

When the user connects a repo, acknowledge it in human terms and store the source so later `show backlog` or `work on #42` requests know where to look.

## Lifecycle Overview

### 1. Issue picked up

An issue enters the Squad flow when one of the following happens:

- the user asks Squad to pull or work the backlog
- Ralph detects an untriaged `squad` issue
- the Lead manually looks at a connected repository backlog

### 2. Triage

The Lead owns triage for base `squad` issues.

Triage actions may include:

- reading the issue
- assigning `squad:{member}` or `squad:copilot`
- commenting with routing notes
- deciding that the issue needs clarification before work starts

### 3. Branch created

Once an issue is routed, the worker creates a branch.

#### Branch naming

| Worker type | Branch pattern |
|-------------|----------------|
| Spawnable squad member | `squad/{issue-number}-{slug}` |
| `@copilot` | `copilot/{issue-number}-{slug}` |

Examples:

```text
squad/42-fix-auth-timeout
copilot/108-add-word-toolbar-test
```

### 4. Work done

The assigned worker implements the change, runs the relevant validation, and commits with the issue reference.

Suggested commit message pattern:

```text
Fix auth timeout handling (#42)
```

### 5. PR opened

Open a PR back to the repo's default integration flow.

Use `gh pr create` when working through the CLI.

### 6. Reviewed

A reviewer evaluates the PR.

Possible outcomes:

- approved
- changes requested
- informational comments only

### 7. Merged

If the PR is approved and checks are green, merge using squash merge.

The repository policy for this project is squash merge only.

## Spawn Prompt Additions for Issue Work

Whenever the coordinator spawns a squad member on issue work, add an `ISSUE CONTEXT` block to the prompt.

### Required block format

```text
ISSUE CONTEXT
Repository: owner/repo
Issue: #42 — Fix auth endpoint timeout
URL: https://github.com/owner/repo/issues/42
Labels: squad, squad:irving, bug
Assignee: Irving
Branch: squad/42-fix-auth-timeout
Acceptance: reduce timeout failures and preserve existing behavior
Linked PR: none yet
Reviewer: Harmony
```

### Rules for the block

1. Always include `Repository` in `owner/repo` format.
2. Include the exact issue number and title.
3. Include the expected branch name if known.
4. Include acceptance criteria or the clearest available success signal.
5. Include the reviewer if one is already obvious.
6. Include the linked PR number if the task is a revision or continuation, otherwise say `none yet`.

### Why it matters

The issue context block keeps the agent anchored to the GitHub artifact rather than treating the task as a free-floating coding request.

## Listing and Presenting Backlog Work

When showing backlog work to the user, use a compact table.

```md
| Issue | Title | Labels | Suggested owner | Notes |
|------:|-------|--------|-----------------|-------|
| #42 | Fix auth endpoint timeout | `squad`, `bug` | Harmony | needs triage |
| #51 | Add Word toolbar coverage | `squad:dylan`, `test` | Dylan | clear scope |
```

Presentation rules:

- show issue number, title, and routing recommendation
- keep it short enough to scan
- call out blockers or ambiguity explicitly

## PR Creation Guidance

Recommended `gh` flow:

```bash
git checkout -b squad/42-fix-auth-timeout
# ... make changes ...
git push -u origin squad/42-fix-auth-timeout
gh pr create --fill --draft
```

### PR body essentials

Every issue-driven PR should include:

- `Closes #42`
- a concise summary of the change
- validation run
- reviewer expectations
- if applicable, a note that the task was routed through `@copilot` under 🟡 review rules

Suggested PR body snippet:

```md
## Summary
- fix timeout handling in the auth proxy path
- preserve existing retry behavior

## Validation
- npm run test:integration

Closes #42
```

## PR Review Handling

### Approved PRs

If the PR is approved and required checks are green:

1. mark it as ready to merge on the board
2. merge using squash merge
3. close or verify closure of the linked issue
4. log the result through normal After Agent Work handling

### Changes requested

If review requests changes:

1. route the revision to the correct next author
2. enforce reviewer rejection lockout if the review is a rejection
3. include the PR number and review summary in the next `ISSUE CONTEXT` block
4. keep the issue and PR linked in status reporting

### Informational comments only

If comments are non-blocking, route them as optional follow-up rather than a locked revision.

## Reviewer Rejection Lockout in the Issue Lifecycle

If a reviewer rejects a PR:

- the author of the rejected version is locked out of the next revision for that artifact
- the coordinator must assign a different eligible member
- if the original author was `@copilot`, the next revision must not be authored by `@copilot`
- if the original author was a squad member, that member may not self-revise

Example note for the next worker:

```text
PR #108 was rejected by Harmony.
Original author Irving is locked out of the next revision.
Revise independently and keep the same issue/PR context in mind.
```

## Merge Commands

Use squash merge.

### Immediate squash merge

```bash
gh pr merge 108 --squash --delete-branch
```

### Auto-merge when checks are still running

```bash
gh pr merge 108 --auto --squash --delete-branch
```

### Important rules

- do not merge directly to `main` by local branch merging
- prefer the repository's protected PR flow
- when asking the user to merge manually on GitHub, tell them to use **Squash and merge**

## Ralph Integration

Ralph uses this lifecycle as the basis for its board logic.

| Board state | Lifecycle meaning |
|-------------|-------------------|
| `Untriaged` | issue picked up but not yet assigned |
| `Assigned` | issue routed, branch not yet visible |
| `In Progress` | branch or draft PR exists |
| `Review Feedback` | PR exists with requested changes |
| `Ready` | PR approved and checks green |
| `Done` | issue closed or PR merged during the session |

## Example End-to-End Flow

```text
1. User: work on issue #42
2. Lead reads #42 and routes it to Irving
3. Irving gets ISSUE CONTEXT with repo, issue, branch, and acceptance
4. Irving creates branch: squad/42-fix-auth-timeout
5. Irving pushes and opens draft PR #108
6. Harmony reviews PR #108
7. Harmony approves after checks pass
8. Ralph merges with gh pr merge 108 --squash --delete-branch
9. Issue #42 closes
```

## Review Checklist

- [ ] connected repo recorded in `team.md` as `owner/repo`
- [ ] spawned issue work includes an `ISSUE CONTEXT` block
- [ ] branch naming matches worker type
- [ ] PR body references the issue
- [ ] review outcomes route correctly, including lockout on rejection
- [ ] merges use squash merge commands
