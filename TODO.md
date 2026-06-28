# miku-xlsx2md Main Repository TODO

This file tracks repository-shape and migration items for the `10 Main Application` side.

Feature and implementation backlog remains in [docs/TODO.md](./docs/TODO.md).

## 2026-05-17 Web Separation Cleanup

- [x] Confirm separated Web App repository: <https://github.com/igapyon/miku-xlsx2md-web>
- [x] Remove Web-owned HTML source and generated HTML from this repository
- [x] Remove Web-owned `lht-cmn/`, browser CSS, browser adapter source, and browser UI test
- [x] Keep product core, CLI, diagnostics, fixtures, and core tests in this repository
- [x] Retarget `npm run build` to regenerate core JavaScript instead of Web HTML
- [x] Update README / CONTRIBUTING / notices for the main application role
- [x] Document and automate formal upstream runtime release assets for `miku-xlsx2md-web`
- [x] Move generated JavaScript to ignored `dist/js/` output

## Runtime Release Asset Contract

Local commands:

- `npm run build:runtime`
- `npm run smoke:runtime`
- `npm run stage:runtime-release`

Local generated files:

- `bundle/miku-xlsx2md-runtime.mjs`
- `bundle/miku-xlsx2md-runtime.json`

GitHub Release assets:

- `miku-xlsx2md-runtime-<version>.mjs`
- `miku-xlsx2md-runtime-<version>.json`

Workflow:

- `.github/workflows/ci.yml`
  - Trigger: pushes to `main` / `devel`, pull requests
  - Responsibility: install dependencies with `npm ci` and run `npm test`
- `.github/workflows/release-runtime-bundle.yml`
  - Trigger: `v*` tag push
  - Responsibility: build/test main app, build runtime bundle, smoke runtime bundle, stage release assets, upload release assets to the matching GitHub Release

## Ownership

Main application repository owns:

- TypeScript product core under `src/ts/`
- ignored generated core JavaScript under `dist/js/`
- Node CLI under `scripts/miku-xlsx2md-cli.mjs`
- Node runtime helper under `scripts/lib/xlsx2md-node-runtime.mjs`
- conversion semantics, diagnostics, fixtures, tests, and implementation docs

Web App repository owns:

- browser UI source and generated Single-file Web App HTML
- `lht-cmn/`
- browser adapter code
- Web smoke tests
- Web release assets
- vendored upstream runtime artifact used by the browser app

## Verification Log

- [x] 2026-05-17 `npm test`
  - Passed: 35 test files, 334 tests
- [x] 2026-05-17 `npm run build:runtime`
  - Generated ignored local runtime files under `bundle/`
- [x] 2026-05-17 `npm run smoke:runtime`
  - Passed for version `1.0.0`
- [x] 2026-05-17 `npm run stage:runtime-release`
  - Staged ignored local release assets under `release-assets/`
