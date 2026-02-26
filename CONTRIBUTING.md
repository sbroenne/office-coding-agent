# Contributing

Thanks for your interest in contributing to Office Coding Agent.

> **⛔ Never push or merge directly to `main`.** All changes must go through a pull request on a feature branch. Branch protection is enforced on GitHub.
>
> **Squash merge only.** When merging a PR on GitHub, always use **"Squash and merge"**. Merge commits and rebase merges are disabled.

## Development Setup

- Install dependencies: `npm install`
- Start the dev server (Copilot proxy + Vite): `npm run dev` (requires GitHub Copilot subscription)
- Sideload in Excel Desktop: `npm run start:desktop:excel`

> **Note:** For full AI functionality you need an active GitHub Copilot subscription and must authenticate with `gh auth login` (or equivalent).

## Before Submitting a PR

Please run:

- `npm run lint`
- `npm run typecheck`
- `npm test`
- `npm run test:ui`

If your change touches Excel host runtime behavior (`Excel.run` paths), also run:

- `npm run test:e2e`

## Contribution Guidelines

- Keep changes focused and minimal.
- Follow existing architecture: single UI, host-routed runtime behavior.
- Add or update tests for any behavior change.
- Avoid introducing unrelated refactors in feature/fix PRs.
- No live API credentials are needed for integration tests — they run in jsdom without a Copilot connection.

## Pull Request Checklist

- [ ] Code compiles and passes checks locally
- [ ] Tests added/updated where appropriate
- [ ] Documentation updated when behavior changes
- [ ] PR description clearly explains what and why

## Reporting Issues

Please include:

- Expected behavior
- Actual behavior
- Reproduction steps
- Environment details (OS, Office host, Node version)

Thanks for contributing.
