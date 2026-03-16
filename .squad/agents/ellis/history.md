# Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook with full Copilot CLI plugin support
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket + JSON-RPC, Vite 7, Vitest, Playwright, Mocha (E2E)
- **Created:** 2026-03-15

## Learnings

- **Squad adoption documented** (2026-01-XX): README.md and CONTRIBUTING.md updated to mention that the project is developed with a Squad AI team. Squad orchestrates collaborative development via named agents configured in `.squad/`. Contributors are directed to `.squad/team.md` for current team composition and roles.
- **Product & doc review completed** (2026-03-16): Comprehensive audit of README, GETTING_STARTED, docs/, agents, tools, and test coverage across all hosts. Key findings: Excel/PowerPoint/Word are solid; Outlook is thin (8 E2E tests vs 233 for Excel, minimal agent definition). No onboarding guidance for new users. Developer docs sparse (no tool API reference, skill/plugin dev guides). Top priorities: strengthen Outlook, add welcome screen, create developer documentation.

<!-- Append new learnings below. Each entry is something lasting about the project. -->
