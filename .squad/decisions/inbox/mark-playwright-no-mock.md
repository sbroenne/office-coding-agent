# Mark: Remove Playwright WebSocket Mocking

## Context
Project policy says Playwright tests in `tests-ui/` must verify real end-to-end behavior through the live proxy and Copilot API. The prior fixture layer used `page.routeWebSocket` plus synthetic JSON-RPC events, which invalidated that coverage boundary.

## Decision
- Remove all Playwright WebSocket mocking helpers from `tests-ui/fixtures.ts`.
- Keep only real-session fixtures (`taskpane`, `configuredTaskpane`) and a lightweight live-server availability helper.
- Rewrite mocked connection/tool-card specs to use the live proxy path.
- Remove the disconnected-state Playwright spec because it depended entirely on a mocked transport; offline/error-path coverage should remain in integration tests unless reproduced without mocks.

## Validation
- `npm run test:integration` was executed.
- Result: failed because live Copilot/proxy integration tests could not connect to `wss://localhost:3000/api/copilot` in this run. Per task instruction, I did not start the dev server manually.

## Follow-up
- PR should note that `tests-ui` now requires the real dev server/proxy path and must not reintroduce `routeWebSocket`.
