# miku-xlsx2md Migration TODO

この `TODO.md` は、`xlsx2md` から `miku-xlsx2md` へ引越しした直後の、リポジトリ名・公開導線・ドキュメント整合の洗い出しメモです。

`docs/TODO.md` は機能開発タスクの正本として既に存在するため、このファイルは移設対応だけを扱います。

## 方針

- 直す対象は、外部公開名、公開 URL、リポジトリ URL、案内文、メタデータ
- すぐには直さない対象は、内部実装の識別子、テスト名、スクリプト名、生成物名、仕様書ファイル名
- つまり `xlsx2md` という文字列が残っていても、内部コード名として妥当なら無理に変えない

## 優先度 A: まず直す

- [ ] GitHub / GitHub Pages の旧 URL を新リポジトリ名へ更新する
  - 現在の旧 URL:
    - `https://github.com/igapyon/xlsx2md`
    - `https://igapyon.github.io/xlsx2md/`
  - 移行先:
    - `https://github.com/igapyon/miku-xlsx2md`
    - `https://igapyon.github.io/miku-xlsx2md/`
  - 修正元:
    - `index-src.html`
    - `miku-xlsx2md-src.html`
  - 生成物:
    - `index.html`
    - `xlsx2md.html`
  - 備考:
    - 生成物ではなく `*-src.html` を修正し、必要なら再ビルドする

- [ ] ランディングページの OGP / Twitter image URL を新公開パスへ更新する
  - 対象:
    - `index-src.html`
    - `index.html`
  - 現状:
    - OGP / Twitter image が旧 `xlsx2md` Pages パスを向いている

- [ ] `CONTRIBUTING.md` のプロジェクト名を新名称に合わせる
  - 現状:
    - `# Contributing to xlsx2md`
    - `` `xlsx2md` へのコントリビュート ``
  - 備考:
    - 文書内容自体は概ねそのままでよいが、リポジトリ移設後の名義だけずれている

## 優先度 B: 早めに判断して直す

- [ ] README に新リポジトリ URL / 公開 URL を明記する
  - 現状:
    - README の表示名は `miku-xlsx2md` に統一済み
    - 新しい GitHub URL や Pages URL の導線も概ね更新済み
  - 方針候補:
    - 冒頭付近に公開先 URL を追加
    - `Use` セクションまたは `What is this?` 付近にリポジトリ URL を追加

- [ ] `package.json` / `package-lock.json` の package 名をどうするか決める
  - 現状:
    - `package.json` の `name` は `miku-xlsx2md` に更新済み
    - `package-lock.json` 側も必要なら追随させる
  - 判断ポイント:
    - 非公開のローカル用メタデータとして維持するなら変更不要
    - `miku-xlsx2md` リポジトリの package として揃えるなら `package-lock.json` も更新する
  - 対象:
    - `package.json`
    - `package-lock.json`

- [ ] `THIRD_PARTY_NOTICES.md` と `CONTRIBUTORS.md` の表記を新名称に寄せるか決める
  - 現状:
    - 本文中に `xlsx2md` が残る
  - 判断ポイント:
    - プロダクト名として `xlsx2md` を残すのか
    - リポジトリ名・配布名として `miku-xlsx2md` を出し分けるのか

## 優先度 C: 必要なら整理

- [x] 公開ファイル名を `miku-xlsx2md.html` に変更する
  - 現状:
    - `README.md`
    - `CONTRIBUTING.md`
    - `index-src.html`
    - `index.html`
    - `scripts/build-miku-xlsx2md.mjs`
    - `miku-xlsx2md-src.html`
  - 判断ポイント:
    - リポジトリ名に合わせて `miku-xlsx2md.html` へ統一する
    - ビルド・README・リンク導線・関連文書をまとめて更新する

- [ ] 仕様書・記事・補助文書のファイル名を rename するか決める
  - 例:
    - `docs/xlsx2md-spec.md`
    - `docs/xlsx2md-impl-spec.md`
    - `docs/xlsx2md-readme-details.md`
    - `docs/articles/qiita/*xlsx2md*`
  - 判断ポイント:
    - 既存参照が多く、変更コストが高い
    - まずは本文や公開導線だけ合わせ、ファイル名は当面維持でもよい

## 直さなくてよさそうなもの

- [ ] 内部コード上の `xlsx2md` 識別子は当面維持する
  - 例:
    - `globalThis.__xlsx2mdModuleRegistry`
    - `requireXlsx2md...`
    - `tests/xlsx2md-*.test.js`
    - `scripts/miku-xlsx2md-cli.mjs`
  - 理由:
    - これはプロダクト内部名として一貫しており、リポジトリ移設だけを理由に触ると差分が大きい

- [ ] 生成物 `index.html` / `miku-xlsx2md.html` を直接編集しない
  - 修正元:
    - `index-src.html`
    - `miku-xlsx2md-src.html`
    - 必要に応じて `scripts/build-miku-xlsx2md.mjs`

## 今回の洗い出しで見つかった代表箇所

- `index-src.html`
  - 旧 GitHub Pages URL
  - 旧 GitHub repository URL
  - OGP / Twitter image URL

- `miku-xlsx2md-src.html`
  - 旧 GitHub repository URL

- `CONTRIBUTING.md`
  - 旧プロジェクト名見出し

- `package.json`
  - package 名は `miku-xlsx2md` に更新済み

- `package-lock.json`
  - package 名が `local-html-tools`

## 補足

- `README.md` の表示名、公開 HTML 名、命名方針は更新済み
- 旧 URL は移行メモとして `TODO.md` にだけ残している
- 外向けの整理は一通り完了しており、残課題は内部名をどこまで維持するかの運用判断が中心
