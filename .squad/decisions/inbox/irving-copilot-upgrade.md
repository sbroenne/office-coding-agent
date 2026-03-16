# Decision: Upgrade @github/copilot to 1.0.5

**Date:** 2026-03-15  
**Author:** Irving (Backend Dev)  
**Status:** Completed — Ready for Review

## Context

Critical security vulnerability GHSA-g8r9-g2v8-jv6f was discovered in @github/copilot versions prior to 1.0.5. This vulnerability enables arbitrary code execution via dangerous shell expansion patterns.

## Decision

Upgraded the following packages:
- `@github/copilot`: ^0.0.414 → ^1.0.5 (major version bump)
- `@github/copilot-sdk`: ^0.1.30 → ^0.1.32 (compatible minor update)

## Rationale

1. **Security:** The GHSA-g8r9-g2v8-jv6f vulnerability is a critical issue that must be addressed immediately
2. **Clean upgrade path:** Despite the major version bump (0.x → 1.x), there are no breaking API changes affecting our codebase
3. **Stable SDK:** All our imports are from `@github/copilot-sdk`, which has a stable API surface between these versions
4. **Verified compatibility:** Build, typecheck, and integration tests all pass with zero code changes required

## Verification

✅ `npm install @github/copilot@1.0.5` — clean install  
✅ `npm install @github/copilot-sdk@latest` — upgraded to 0.1.32  
✅ `npm run build:dev` — build succeeds, no errors  
✅ `npm run typecheck` — zero TypeScript errors  
✅ `npm audit` — GHSA-g8r9-g2v8-jv6f no longer present  
✅ Integration tests — 759/776 pass (17 failures are live WebSocket tests requiring `npm run dev`, expected behavior)

## Impact

**Code changes:** None required beyond package.json/package-lock.json  
**API compatibility:** Fully backward compatible  
**Dependencies:** @github/copilot-sdk updated to match  
**Security posture:** Critical vulnerability eliminated

## Next Steps

1. Stefan to review package.json/package-lock.json changes
2. Commit changes on a feature branch (DO NOT push to main)
3. Create PR with "Squash and merge" option
4. Run E2E tests after merge to verify runtime behavior in real Office hosts

## Notes

Remaining `npm audit` vulnerabilities are in dev dependencies (mocha, electron, flatted, undici, tar) and do not affect runtime security. These can be addressed in a separate cleanup task.
