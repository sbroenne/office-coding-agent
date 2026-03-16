# Squad Team

> office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook

## Coordinator

| Name | Role | Notes |
|------|------|-------|
| Squad | Coordinator | Routes work, enforces handoffs and reviewer gates. |

## Members

| Name | Role | Charter | Status |
|------|------|---------|--------|
| Harmony | Lead | `.squad/agents/harmony/charter.md` | 🏗️ Active |
| Ellis | Product Manager | `.squad/agents/ellis/charter.md` | 📋 Active |
| Dylan | Frontend Dev | `.squad/agents/dylan/charter.md` | ⚛️ Active |
| Irving | Backend Dev | `.squad/agents/irving/charter.md` | 🔧 Active |
| Mark | Tester | `.squad/agents/mark/charter.md` | 🧪 Active |
| Parker | UX Tester | `.squad/agents/parker/charter.md` | 🧪 Active |
| Scribe | Session Logger | `.squad/agents/scribe/charter.md` | 📋 Active |
| Ralph | Work Monitor | — | 🔄 Monitor |
| @copilot | Coding Agent | `copilot-instructions.md` | 🤖 Active |
<!-- copilot-auto-assign: true -->

## @copilot Capability Profile

| Category | Fit | Notes |
|----------|-----|-------|
| Bug fixes with repro steps | 🟢 | Well-defined, bounded scope |
| Test additions (existing patterns) | 🟢 | Follows established conventions |
| Dependency updates | 🟢 | Mechanical, low judgment |
| Documentation updates | 🟢 | Low risk, clear scope |
| Small feature (clear spec, existing patterns) | 🟡 | Needs PR review by Lead |
| Refactoring with test coverage | 🟡 | Needs PR review by Lead |
| New tool definitions | 🟡 | Pattern exists but host-specific nuance |
| Architecture / design decisions | 🔴 | Requires team judgment |
| Security-sensitive changes | 🔴 | Auth, credentials, access control |
| VS Code design system compliance | 🔴 | Requires visual judgment |
| Cross-host runtime behavior | 🔴 | Requires Office host expertise |
| New agent/skill definitions | 🔴 | Requires prompt engineering expertise |

## Issue Source

- **Repository:** sbroenne/office-coding-agent
- **Connected:** 2026-03-15
- **Filters:** All open issues with `squad` label
- **Auto-assign @copilot:** Enabled

## Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in with GitHub Copilot, CLI plugin support, host-routed tools
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket, Vite 7, Vitest, Playwright, Mocha
- **Created:** 2026-03-15
