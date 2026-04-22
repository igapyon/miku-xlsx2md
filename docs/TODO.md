# TODO xlsx2md

この文書は「まだやること全部」を 1 か所に置くためのメモだが、現時点では

- 今すぐ閉じるべき整合不良
- 実装・設計として継続する中長期バックログ

が混在していた。以後はこの 2 つを分けて扱う。

## 今すぐ閉じるべき整合不良

- 文書の現行実装追随を優先する
  - `README.md`、`docs/xlsx2md-readme-details.md`、`docs/xlsx2md-impl-spec.md`、`docs/xlsx2md-spec.md` の記述差を減らす
  - 特に「実装済みなのに未実装と読める記述」を先に潰す
- `docs/xlsx2md-impl-spec.md` の encoding / BOM 記述を現行コードへ合わせる
  - 現在は GUI / CLI / ZIP で `encoding` / `bom` を扱える
  - 古い「UTF-8 固定」前提の説明を除去または更新する
- `docs/TODO.md` から完了済み項目を backlog から外す
  - `border` table detection mode の追加は完了済み
  - `section block` は no-op ではなく軽量 grouping として実装済み
  - `Markdown 保存時の UTF-8 BOM 方針` は「未着手」ではなく「既定値や文書整合の見直し」へ格下げする
- `進捗メモ` の古い表現を更新する
  - `core.ts` の no-op `extractSectionBlocks(...)` という表現は現状に合っていない

## 中長期バックログ

### 近い将来の優先タスク

- レイアウト中心シートの分解改善
  - 根拠:
    - `docs/local-data-review.md` で `イベント プランナー`、`月間プランナー`、`老後資金プランナー` の差分が継続課題として残っている
    - 現在の軽量 section grouping はあるが、巨大表化や小表への分裂をまだ十分に抑えきれていない
  - 当面の対象:
    - `TFc2b640.../イベント プランナー`
    - `TF97739.../月間プランナー`
    - `TF0ffdef.../老後資金プランナー`
  - 狙い:
    - 巨大表 1 個や曜日ごとの小表乱立を減らし、`セクション / 表 / リスト / 画像` のまとまりを自然に出す

- formula の AST evaluator 適用範囲整理
  - 根拠:
    - 手元実データの観測では AST parse の穴埋めよりも evaluator 適用順と表示差分確認が次段になっている
    - `cached value -> AST evaluator -> 既存 resolver -> fallback_formula` の順は固まっている
  - 狙い:
    - 実データ頻出関数を AST evaluator 側へ寄せつつ、既存 resolver 依存を観測しながら縮小する

- Markdown escape 方針の統一
  - 根拠:
    - `plain / github` 分離や token 化は進んでいるが、narrative / heading / list / table の文脈差がまだ散っている
  - 狙い:
    - 出力の一貫性を上げ、回帰テストの見通しを良くする

- `docs/local-data-review.md` に基づく実データレビュー継続
  - 人手確認が効く対象がはっきりしている
  - 実装改善と文書方針確認の両方に直接効く

### 優先度中

- formula 次段タスク
  - `scripts/observe-xlsx2md-formulas.mjs` による観測を継続し、AST evaluator 側へ寄せる関数群を整理する
  - 優先順は `cached value -> AST evaluator -> 既存 resolver -> fallback_formula` で固定
  - 既存の文字列ベース evaluator は互換 fallback として維持しつつ、AST evaluator 側へ段階移行する
  - 既存 resolver は現時点では安全装置として必要であり、短期的には削除しない
  - 中長期的には、実データ観測を踏まえて担当範囲を縮小し、後方互換 fallback へ寄せる
  - 次候補:
    - `XLOOKUP` の binary search `search_mode=2/-2` の境界条件を必要に応じて詰める
    - 実データ頻出関数を AST evaluator 側へさらに寄せる
- `セクション分割ブロック` の次段整理
  - 現在の軽量 grouping を、専用の意味分解器へ進めるか判断する
- UI の `formulaDiagnostics` / `tableScores` 表示を見直す
- 表検出モードの次段整理
  - `balanced / border` は実装済み
  - 次は mode 差分の説明、fixture、実データ適用方針を整理する
  - 設計メモ: `docs/table-detection-border-priority.md`
- Markdown エスケープを統一する
  - 表セル、narrative、見出し、箇条書きで共通方針を持つ
  - 少なくとも `改行 / | / \`` を安全に扱う
  - 必要に応じて行頭の Markdown 記号 (`#`, `-`, `*`, `>`) も整理する
  - `github` formatting mode ではセル内改行を `<br>` として扱う方針に寄せた
  - `markdown escape -> rich text parser -> plain/github formatter -> table escape` の段階分離までは実装済み
  - 次は `escaped` part と table / narrative / heading / list それぞれの escape 方針をどこまで共通化するかを整理する
  - 小さい追加テスト候補:
    - normalize 層での table cell 空白/タブの期待整理
    - parser / formatter / renderer の対称比較の見出し整理
    - fixture README と個別 test 名の対応づけ整理
  - narrative / heading / list item の文脈差は未整理
  - GitHub 以外の Markdown viewer 差は未整理
  - 空白 / 改行 / run 境界 / 装飾入れ子の境界ケースは追加検証余地あり
  - 詳細メモは `docs/rich-text-markdown-rendering.md` を参照

- rich text / Markdown renderer の責務整理
  - 現状は `shared-strings.ts` / `styles-parser.ts` / `worksheet-parser.ts` / `sheet-markdown.ts` に加えて、`markdown-escape.ts` / `rich-text-parser.ts` / `rich-text-plain-formatter.ts` / `rich-text-github-formatter.ts` / `markdown-table-escape.ts` へ段階分離済み
  - `plain` と `github` の責務境界は formatter 分離でかなり明確になった
  - `sheet-markdown.ts` は、セル文字列化 / narrative item 構築 / section block 判定 / grouped body 組み立て / asset section 生成 / render state 収集 / Markdown summary 組み立てまで分離済み
  - `sheet-markdown.ts` の整理に合わせて `tests/xlsx2md-sheet-markdown.test.js` に直接テストを追加し、回帰確認しやすくした
  - 必要なら token ベースの中間表現をさらに細粒度化する
  - `styledText.parts` は `text / escaped` と `rawText` を持つ段階まで進めたので、次はそれを renderer policy にどう使うかを詰める
  - GUI / CLI の既定値差は今後も注意して確認する
  - rich fixture 回帰:
    - `tests/fixtures/rich/rich-text-github-sample01.xlsx`
    - `tests/fixtures/rich/rich-markdown-escape-sample01.xlsx`
- 共有型定義の単一正本を作る
  - `ParsedCell` / `ParsedSheet` / `ParsedWorkbook` / `MarkdownFile` などが `core.ts` / `sheet-markdown.ts` / `formula-resolver.ts` / `module-registry-access.ts` などに重複している
  - 少なくとも workbook / sheet / formula / markdown export の主要型は共有 `types` へ寄せたい
  - 型変更時に 1 箇所だけ更新して別モジュールが古い前提のまま残る状態を減らしたい
- `module-registry-access.ts` の巨大な手書き契約を縮小する
  - 現状は module registry 越しの API 型を access 層で大きく再定義している
  - 各モジュールの公開 interface を小さく保ち、access 側の重複定義を減らしたい
  - 文字列名による結合をすぐにやめなくても、契約の単一正本は作りたい
- `src/ts` と `src/js` の二重管理コストを下げる
  - 現状は `src/ts` を正本としつつ、生成済み `src/js` もリポジトリに載せている
  - レビュー差分が大きくなりやすく、変更コストも高い
  - build / test / 配布の都合を整理したうえで、どこまで生成物を持つかを見直したい
- `core.ts` の責務をさらに分割する
  - 現状でも改善は進んでいるが、型定義・変換フロー・export 周辺の集積点としてはまだ太い
  - 「中心に何でも集まるファイル」から、薄い orchestration へ寄せたい
  - 少なくとも workbook model / conversion pipeline / export composition の境界はより明確にしたい
- formula 系の `評価 / 状態更新 / fallback 制御` をさらに分離する
  - 現状は AST evaluator / legacy resolver / cached value 優先の方針はあるが、実装では層がまだ混ざる
  - 例外ベースの未解決制御や、resolver 中での破壊的な cell 更新の境界は整理余地がある
  - 中長期的には evaluator の返り値と state mutation を分離し、fallback policy を見通しよくしたい
- Markdown 保存まわりの既定値・説明整理
  - 対象経路: GUI の `Save Markdown`、CLI の `--out`、ZIP 内 Markdown
  - `encoding` / `bom` 自体は実装済み
  - 残件は、既定値の妥当性と README / spec / impl-spec / 回帰テストの説明整合である

- merge multiline 次段タスク
  - `tests/fixtures/merge/merge-multiline-sample01.xlsx` は追加済み
  - 結合セル内改行の Markdown 正規化ポリシー変更を行う場合は、`markdown-normalize.ts` と `sheet-markdown.ts` と回帰テストをセットで見直す

### 長期バックログ

- Qiita 記事の新規作成を検討する
  - 直近のハイパーリンク対応、GitHub 向け Markdown 出力方針、fixture hygiene (`x15ac:absPath`) の知見を記事化したい
  - 実装断片だけでなく、なぜその出力方針にしたかも整理して残したい
  - 候補は `docs/articles/qiita/` 配下へ新規記事を追加
  - 既存記事と重複しにくい題材候補:
    - rich text / Markdown escape / GitHub formatting mode の設計と割り切り
    - `balanced / border` の table detection mode 差分と使い分け
    - shape / chart を「見た目再現」ではなく意味情報として Markdown 化する方針

## 未対応事項

- 数式未対応の整理
  - `space intersection` の完全対応
  - 配列定数の完全対応
  - dynamic array / spill の完全対応
  - `LAMBDA / LET / MAP / REDUCE / SCAN`
  - 完全な `R1C1` 文法
  - Excel の future function 全般
  - `NOW` など volatile 関数の完全再計算
- レイアウト未対応の整理
  - `セクション分割ブロック` の導入
  - `カレンダー / ボード / ダッシュボード系` シートの専用扱い
  - レイアウト中心シートの完全再現は対象外であり、`セクション / 表 / リスト / 画像` 分解で扱う
  - `イベント プランナー` のようなフォーム風罫線領域は、現時点では保守的に扱う
  - DrawingML の図形 (`xdr:sp` / `xdr:cxnSp` など) は、現時点では安全に無視またはメタデータ抽出に留める
  - `DrawingML -> SVG` は将来候補
  - グラフは当面、意味情報のテキスト化で固定し、`Chart -> SVG` は保留とする
  - SmartArt は現時点では fallback とし、意味解釈や SVG 化の対象外とする
- ハイパーリンク次段整理
  - 外部 URL とブック内リンクの Markdown 出力は実装済み
  - ブック内リンクは現時点では対象シート先頭アンカーへのリンクを基本とする
  - Excel ブック内リンクの `sheet / cell / range` 追跡を、必要ならより厳密にする
  - hyperlink セルは GitHub 出力で underline を重ねて出さない方針
  - rich text と共存する場合の境界ケースは追加確認余地がある

## 方針未確定

- `XLOOKUP` 近似一致や binary search を未ソート範囲でどう扱うか
- `ROW / COLUMN` の文脈なし引数なし形をどう扱うか
- 配列定数をどこまで Excel 互換で広げるか
- `A1#` のような spill 演算子を、runtime でどこまで実解決するか
- `f@ref` を spill と array formula でどう見分けるか
- `TODAY / NOW` を cached value 専用に留めるか
- `existing resolver` から AST evaluator へ、どこまで段階移行したら縮小判断するか

## レイアウト系の整理

- 手元実データの差分レビュー継続
  - 重点対象は `docs/local-data-review.md` を参照
  - 優先順:
    - `TFc2b640.../イベント プランナー` の多画像・多結合差分確認
    - `TF97739.../月間プランナー` の merge 多用差分確認
  - 人手確認があると助かるもの:
    - 代表シートの Excel スクリーンショット
    - 条件付き書式やアイコン列の意味説明
- レイアウト中心シート方針の維持と具体化
  - 見た目再現ではなく、`セクション / 表 / リスト / 画像` への分解を優先する
- `セクション分割ブロック` 導入検討
  - 対象候補: 入力パネル、概要カード、見出し付きの広い merge 領域
- `カレンダー / ボード / ダッシュボード系` シートの別カテゴリ化検討
  - 対象候補: `TF97739.../月間プランナー`
  - 対象候補: `TFc2b640.../イベント プランナー`
  - 対象候補: `TF0ffdef.../老後資金プランナー`

## 進捗メモ

- fixture ベースの実ファイル調整は一段落
- formula Step 2 の最小 parser 土台は追加済み
- formula Step 3 の最小 evaluator 土台は追加済み
- `extractSectionBlocks(...)` は no-op ではなく、縦ギャップに基づく軽量 section grouping として実装済み
- `sheet-markdown.ts` の責務整理を実施し、セル整形 / narrative / section grouping / asset section / render state / summary の分割を進めた
- `tests/xlsx2md-sheet-markdown.test.js` に回帰テストを追加し、`sheet-markdown` 単体テストは 23 件通過を確認済み
- `balanced / border` の table detection mode は GUI / CLI と回帰テストまで含めて実装済み
- Markdown の `encoding` / `bom` 切替は GUI / CLI / ZIP 出力まで実装済み

## 参照

- 正本: `docs/TODO.md`
- 関連仕様:
  - `docs/xlsx2md-spec.md`
  - `docs/xlsx2md-impl-spec.md`
  - `docs/xlsx-formula-subset.md`
