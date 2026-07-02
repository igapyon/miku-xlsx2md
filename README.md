# miku-xlsx2md

`miku-xlsx2md` is the TypeScript / Node.js main application for converting Excel (`.xlsx`) workbooks into Markdown-oriented artifacts.

This repository owns the product core, Node.js CLI, conversion semantics, diagnostics, fixtures, and tests. The separated browser Web App surface is maintained in `miku-xlsx2md-web`.

Links:

- Main application repository: <https://github.com/igapyon/miku-xlsx2md>
- Web App repository: <https://github.com/igapyon/miku-xlsx2md-web>
- Web App: <https://igapyon.github.io/miku-xlsx2md-web/>

## What is this?

`miku-xlsx2md` reads `.xlsx` files locally and extracts prose, tables, images, chart information, shape source data, comments, notes, threaded comments, hyperlinks, rich text, and formula-derived values into Markdown and related artifacts.

The conversion goal is meaningful Markdown extraction, not exact visual reproduction of Excel.

Generated combined Markdown includes workbook-level YAML front matter with the artifact title, type, and conversion settings.

## Features

- Converts all sheets in a workbook in one pass
- Extracts prose, tables, images, chart configuration data, shape source data, comments, notes, and threaded comments
- Detects table-like regions using borders and value groupings
- Supports `display / raw / both` output modes
- Supports `plain / github` formatting modes
- Supports `balanced / border / planner-aware` table detection modes
- Adds workbook-level YAML front matter to generated combined Markdown
- Preserves supported rich text and hyperlinks where practical
- Prefers cached formula values and parses formulas when needed
- Writes Markdown or ZIP output from the Node.js CLI
- Provides a runtime helper used by downstream surfaces such as `miku-xlsx2md-web`

## Node CLI

The CLI converts one input workbook at a time and writes Markdown or ZIP output to a file.

```bash
npm run cli -- ./tests/fixtures/xlsx2md-basic-sample01.xlsx --out /tmp/xlsx2md-basic.md
```

```bash
npm run cli -- ./tests/fixtures/xlsx2md-basic-sample01.xlsx --zip /tmp/xlsx2md-basic.zip
```

Options:

- `--out <file>`: Write combined Markdown to a file
- `--zip <file>`: Write ZIP export to a file
- `--output-mode <mode>`: `display`, `raw`, or `both`
- `--formatting-mode <mode>`: `plain` or `github`
- `--table-detection-mode <mode>`: `balanced`, `border`, or `planner-aware`
- `--encoding <value>`: `utf-8`, `shift_jis`, `utf-16le`, `utf-16be`, `utf-32le`, or `utf-32be`
- `--bom <value>`: `off` or `on`
- `--shape-details <mode>`: `include` or `exclude`
- `--front-matter <mode>`: `include` or `exclude`
- `--include-shape-details`: Alias for `--shape-details include`
- `--no-header-row`: Do not treat the first row as a table header
- `--no-trim-text`: Preserve surrounding whitespace
- `--keep-empty-rows`: Keep empty rows
- `--keep-empty-columns`: Keep empty columns
- `--summary`: Print per-sheet summary to stdout
- `--help`: Show help and exit

Output contract for agents:

- Primary output is one workbook-level combined Markdown document.
- Combined Markdown starts with YAML front matter by default. Use `--front-matter exclude` to omit it.
- The Markdown body starts with `# Book: <workbook>` and then `## Sheet: <sheet>` sections in workbook sheet order.
- ZIP output stores the combined Markdown at `output/<workbook>.md` and extracted assets under `output/assets/`.
- Front matter fields are `title`, `type`, and `conversion`.
- `conversion` records the tool version, output mode, formatting mode, table detection mode, and shape detail mode.

Exit codes:

- `0`: Success
- `1`: Error

## Build And Test

```bash
npm install
npm run build
```

`npm run build` transpiles the TypeScript product core into `dist/js/` through the test pre-step and then runs the test suite. `src/ts/` is the source of truth; `dist/js/` is ignored generated output used by tests, the CLI runtime helper, and runtime bundle generation.

Generated browser HTML files are no longer owned by this repository. Build and release the browser app from `miku-xlsx2md-web`.

GitHub Actions:

- `.github/workflows/ci.yml` runs `npm ci` and `npm test` on pushes to `main` / `devel` and on pull requests.
- `.github/workflows/release-cli-runtime-bundles.yml` runs when a `v*` GitHub Release is published and publishes CLI/runtime/source release assets.

## Runtime Bundle

This repository publishes the upstream runtime bundle consumed by downstream surfaces such as `miku-xlsx2md-web`.

```bash
npm run build:bundle
npm run smoke:bundle
npm run smoke:runtime
npm run stage:runtime-release
```

Generated local artifacts:

- `bundle/miku-xlsx2md.mjs`
- `bundle/miku-xlsx2md-runtime.mjs`
- `bundle/miku-xlsx2md-runtime.json`
- `bundle/miku-xlsx2md-sources.tgz`

Release asset names:

- `miku-xlsx2md-<version>.mjs`
- `miku-xlsx2md-runtime-<version>.mjs`
- `miku-xlsx2md-sources-<version>.tgz`

`bundle/miku-xlsx2md-runtime.json` is local build metadata and is not uploaded as a GitHub Release asset.

The GitHub Actions workflow `.github/workflows/release-cli-runtime-bundles.yml` runs when a `v*` GitHub Release is published, checks out the release tag, builds and tests the main application, builds the CLI/runtime/source bundles, runs bundle smoke checks, stages release assets, and uploads them to the matching GitHub Release.

## Tech Stack

- Runtime: Node.js
- Source language: TypeScript
- CLI: `scripts/miku-xlsx2md-cli.mjs`
- Runtime helper: `scripts/lib/xlsx2md-node-runtime.mjs`
- Testing: Vitest and jsdom

## Naming

- Main application repository: `miku-xlsx2md`
- Web App repository: `miku-xlsx2md-web`
- Product / internal name: `xlsx2md`
- Display name: `miku-xlsx2md`

Internal identifiers, script names, tests, fixtures, and specification documents may continue to use `xlsx2md` where that is the stable product-internal name.

## Documentation

- High-level specification and design policy: [docs/xlsx2md-spec.md](./docs/xlsx2md-spec.md)
- Detailed implementation-oriented specification: [docs/xlsx2md-impl-spec.md](./docs/xlsx2md-impl-spec.md)
- Generated Markdown front matter contract: [docs/xlsx2md-front-matter.md](./docs/xlsx2md-front-matter.md)
- miku-soft maintenance reference: [docs/miku-soft-reference.md](./docs/miku-soft-reference.md)
- Development backlog: [docs/TODO.md](./docs/TODO.md)

## License

Released under the Apache License 2.0.

See [LICENSE](./LICENSE) and [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md).

--------------------------------------------------------------------------------

# miku-xlsx2md

`miku-xlsx2md` は、Excel (`.xlsx`) ブックを Markdown 向けの成果物へ変換する TypeScript / Node.js の main application です。

このリポジトリは product core、Node.js CLI、変換仕様、診断情報、fixture、テストを所有します。分離済みのブラウザ Web App は `miku-xlsx2md-web` で管理します。

リンク:

- Main application repository: <https://github.com/igapyon/miku-xlsx2md>
- Web App repository: <https://github.com/igapyon/miku-xlsx2md-web>
- Web App: <https://igapyon.github.io/miku-xlsx2md-web/>

## これは何か

`miku-xlsx2md` は `.xlsx` ファイルをローカルで読み込み、地の文、表、画像、グラフ情報、図形元データ、リンク、rich text、数式由来の値を Markdown と関連成果物として抽出します。

目的は Excel の見た目を完全再現することではなく、意味のある Markdown として情報を取り出すことです。

生成される連結 Markdown には、成果物 title、type、変換設定を示す Workbook 単位の YAML front matter が付きます。

## Node CLI

```bash
npm run cli -- ./tests/fixtures/xlsx2md-basic-sample01.xlsx --out /tmp/xlsx2md-basic.md
```

```bash
npm run cli -- ./tests/fixtures/xlsx2md-basic-sample01.xlsx --zip /tmp/xlsx2md-basic.zip
```

主なオプションは `--output-mode`、`--formatting-mode`、`--table-detection-mode`、`--encoding`、`--bom`、`--shape-details`、`--front-matter`、`--summary` です。詳細は `--help` を参照してください。

生成 AI / agent 向けの出力契約:

- 主出力は Workbook 単位の連結 Markdown です。
- 連結 Markdown は既定で YAML front matter から始まります。`--front-matter exclude` で省略できます。
- Markdown 本文は `# Book: <workbook>` で始まり、Workbook 内のシート順に `## Sheet: <sheet>` セクションが続きます。
- ZIP 出力では、連結 Markdown は `output/<workbook>.md`、抽出 asset は `output/assets/` 配下に入ります。
- front matter のフィールドは `title`、`type`、`conversion` です。
- `conversion` には tool version、output mode、formatting mode、table detection mode、shape detail mode が入ります。

## ビルドとテスト

```bash
npm install
npm run build
```

`npm run build` はテスト前処理で TypeScript の product core を `dist/js/` へ変換し、その後テストを実行します。`src/ts/` が正本であり、`dist/js/` はテスト、CLI runtime helper、runtime bundle 生成が利用する Git 管理外の生成物です。

ブラウザ向け HTML 生成物はこのリポジトリの所有物ではありません。Web App のビルドとリリースは `miku-xlsx2md-web` で行います。

## Runtime Bundle

このリポジトリは、`miku-xlsx2md-web` などの downstream surface が利用する upstream runtime bundle を提供します。

```bash
npm run build:bundle
npm run smoke:bundle
npm run smoke:runtime
npm run stage:runtime-release
```

ローカル生成物:

- `bundle/miku-xlsx2md.mjs`
- `bundle/miku-xlsx2md-runtime.mjs`
- `bundle/miku-xlsx2md-runtime.json`
- `bundle/miku-xlsx2md-sources.tgz`

GitHub Release asset 名:

- `miku-xlsx2md-<version>.mjs`
- `miku-xlsx2md-runtime-<version>.mjs`
- `miku-xlsx2md-sources-<version>.tgz`

`bundle/miku-xlsx2md-runtime.json` はローカル build metadata であり、GitHub Release asset としては upload しません。

`.github/workflows/release-cli-runtime-bundles.yml` は `v*` GitHub Release の publish で動作し、release tag の checkout、main application の build/test、CLI/runtime/source bundle 生成、bundle smoke、release asset staging、GitHub Release への upload を行います。
