# Contributing

Thanks for your interest in contributing to Office Coding Agent.

## Development Setup

- Install dependencies: `npm install`
- Start dev server: `npm run dev`
- Sideload in Excel Desktop: `npm run start:desktop`

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
