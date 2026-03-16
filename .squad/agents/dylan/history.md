# Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook with full Copilot CLI plugin support
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket + JSON-RPC, Vite 7, Vitest, Playwright, Mocha (E2E)
- **Created:** 2026-03-15

## Learnings

<!-- Append new learnings below. Each entry is something lasting about the project. -->
- 2026-03-16: `src\taskpane\App.tsx` still breaks the VS Code-only UI rule by importing `lucide-react` and using amber/destructive utility styling instead of codicons and explicit `--vscode-*` token-driven visuals.
- 2026-03-16: `src\components\ModelPicker.tsx` mounts `useOfficeChat(host)` directly, creating a second chat session lifecycle inside a picker instead of consuming the already-active chat controller.
- 2026-03-16: Several chat controls rely on undeclared VS Code variables (`--vscode-icon-foreground`, badge, keybinding, quick input, widget shadow), so parts of the UI currently depend on browser fallbacks rather than a fully defined VS Code token surface.
- 2026-03-16: `ChatPanel` wires ActionBar regenerate/feedback callbacks to TODO no-ops, so the UI advertises controls that currently do nothing.
- 2026-03-16: `ModelPicker` must stay presentation-only: it reads model options from `useSettingsStore`, but the active session state and `switchModel` callback must be passed down from the single `useOfficeChat` owner in `App`/`ChatPanel` to avoid duplicate WebSocket sessions.
- 2026-03-16: App-level status and permission banners should use codicons plus `--vscode-*` token styling, not standalone icon libraries or host-agnostic warning/error utility classes.
