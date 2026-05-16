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
- [ ] Publish or document a formal upstream runtime release asset for `miku-xlsx2md-web`
- [ ] Revisit whether `src/js/` should remain tracked or move to an ignored runtime output directory after the downstream runtime asset contract is finalized

## Ownership

Main application repository owns:

- TypeScript product core under `src/ts/`
- generated core JavaScript under `src/js/` while the current Node/runtime contract depends on it
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
