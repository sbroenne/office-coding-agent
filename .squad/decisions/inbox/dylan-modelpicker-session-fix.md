# Dylan: ModelPicker Session Ownership Fix

## Decision
`ModelPicker` should not call `useOfficeChat()` directly. It remains a presentational picker backed by `useSettingsStore` for model list/current selection, while `ReadyAssistant` owns the single chat session and passes `hasActiveSession` plus `onSwitchModel` through `ChatPanel`.

## Why
Mounting `useOfficeChat()` inside the picker created a second WebSocket/session lifecycle that could drift from the visible conversation. Passing session-aware props keeps model switching aligned with the active session while preserving no-session behavior (store-only model updates for the next conversation).

## UI note
App-level connection, error, and permission banners were also normalized to codicons and `--vscode-*` tokens so the task pane matches VS Code Copilot Chat more closely.
