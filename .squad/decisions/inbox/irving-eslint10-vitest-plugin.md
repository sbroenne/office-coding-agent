# Irving Decision Inbox — ESLint 10 Vitest plugin

- **Date:** 2026-03-15
- **Owner:** Irving
- **Status:** Proposed

## Context
Upgrading `eslint` / `@eslint/js` to v10 caused the existing `eslint-plugin-vitest` package to fail at runtime. The latest published `eslint-plugin-vitest` release (`0.5.4`) only declares support for ESLint 8/9, and linting crashed with `Class extends value undefined` under ESLint 10.

## Decision
Use `@vitest/eslint-plugin` instead of `eslint-plugin-vitest` in `eslint.config.mjs`.

## Rationale
- `@vitest/eslint-plugin` declares `eslint >=8.57.0`, which covers ESLint 10.
- It preserves the same flat-config usage pattern (`vitest.configs.recommended.rules`, `vitest.environments.env.globals`) with minimal config churn.
- This keeps the repo on the requested ESLint 10 line without pinning lint tooling back to ESLint 9.

## Impact
- `package.json` / `package-lock.json` now depend on `@vitest/eslint-plugin`.
- Future lint-stack upgrades should treat `@vitest/eslint-plugin` as the supported Vitest ESLint integration for this repo.
