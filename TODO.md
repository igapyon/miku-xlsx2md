# miku-xlsx2md Migration TODO

この `TODO.md` は、`xlsx2md` から `miku-xlsx2md` へ引越しした直後の、リポジトリ名・公開導線・ドキュメント整合の洗い出しメモです。

`docs/TODO.md` は機能開発タスクの正本として既に存在するため、このファイルは移設対応だけを扱います。

## 方針

- 直す対象は、外部公開名、公開 URL、リポジトリ URL、案内文、メタデータ
- すぐには直さない対象は、内部実装の識別子、テスト名、スクリプト名、生成物名、仕様書ファイル名
- つまり `xlsx2md` という文字列が残っていても、内部コード名として妥当なら無理に変えない

## 優先度 A: まず直す

- [x] GitHub / GitHub Pages の旧 URL を新リポジトリ名へ更新する
  - 移行先:
    - `https://github.com/igapyon/miku-xlsx2md`
    - `https://igapyon.github.io/miku-xlsx2md/`
  - 修正元:
    - `index-src.html`
    - `miku-xlsx2md-src.html`
  - 生成物:
    - `index.html`
    - `miku-xlsx2md.html`
  - 備考:
    - `*-src.html` と生成物は新 URL へ更新済み

- [x] ランディングページの OGP / Twitter image URL を新公開パスへ更新する
  - 対象:
    - `index-src.html`
    - `index.html`
  - 対応済み:
    - OGP / Twitter image は `https://igapyon.github.io/miku-xlsx2md/docs/screenshots/xlsx2md-ogp.png` を向いている

- [x] `CONTRIBUTING.md` のプロジェクト名を新名称に合わせる
  - 対応済み:
    - `# Contributing to miku-xlsx2md`
    - `` `miku-xlsx2md` へのコントリビュート ``
  - 備考:
    - 文書内容自体は概ねそのままでよいが、リポジトリ移設後の名義だけずれている

## 優先度 B: 早めに判断して直す

- [x] README に新リポジトリ URL / 公開 URL を明記する
  - 対応済み:
    - 冒頭付近に Web app URL と Repository URL を追加

- [x] `package.json` / `package-lock.json` の package 名をどうするか決める
  - 対応済み:
    - `package.json` の `name` は `miku-xlsx2md` に更新済み
    - `package-lock.json` の root `name` も `miku-xlsx2md` に更新済み
  - 対象:
    - `package.json`
    - `package-lock.json`

- [x] `THIRD_PARTY_NOTICES.md` と `CONTRIBUTORS.md` の表記を新名称に寄せるか決める
  - 対応済み:
    - 外向けプロジェクト名としての本文表記は `miku-xlsx2md` へ更新
    - 内部名・仕様書名・fixture 名としての `xlsx2md` は維持

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

- [x] 仕様書・記事・補助文書のファイル名を rename するか決める
  - 例:
    - `docs/xlsx2md-spec.md`
    - `docs/xlsx2md-impl-spec.md`
    - `docs/xlsx2md-readme-details.md`
    - `docs/articles/qiita/*xlsx2md*`
  - 対応方針:
    - 既存参照が多く、変更コストが高いため当面 rename しない
    - README の命名方針どおり、仕様書名・記事名・補助文書名には `xlsx2md` が残ってよい

## 直さなくてよさそうなもの

- [x] 内部コード上の `xlsx2md` 識別子は当面維持する
  - 例:
    - `globalThis.__xlsx2mdModuleRegistry`
    - `requireXlsx2md...`
    - `tests/xlsx2md-*.test.js`
    - `scripts/miku-xlsx2md-cli.mjs`
  - 理由:
    - これはプロダクト内部名として一貫しており、リポジトリ移設だけを理由に触ると差分が大きい

- [x] 生成物 `index.html` / `miku-xlsx2md.html` を直接編集しない
  - 修正元:
    - `index-src.html`
    - `miku-xlsx2md-src.html`
    - 必要に応じて `scripts/build-miku-xlsx2md.mjs`
  - 対応方針:
    - 生成物は build で更新する

## 今回の洗い出しで見つかった代表箇所

- `index-src.html`
  - GitHub Pages URL は更新済み
  - GitHub repository URL は更新済み
  - OGP / Twitter image URL は更新済み

- `miku-xlsx2md-src.html`
  - GitHub repository URL は更新済み

- `CONTRIBUTING.md`
  - プロジェクト名見出しは更新済み

- `package.json`
  - package 名は `miku-xlsx2md` に更新済み

- `package-lock.json`
  - package 名は `miku-xlsx2md` に更新済み

## 補足

- `README.md` の表示名、公開 HTML 名、命名方針は更新済み
- 旧 GitHub / GitHub Pages URL は外向け文書から除去済み
- 外向けの整理は一通り完了しており、内部名としての `xlsx2md` は当面維持する
