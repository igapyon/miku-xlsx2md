# TODO xlsx2md

## 実装タスク

- formula 次段タスク
  - `scripts/observe-xlsx2md-formulas.mjs` による観測を継続し、AST evaluator 側へ寄せる関数群を整理する
  - 優先順は `cached value -> AST evaluator -> 既存 resolver -> fallback_formula` で固定
  - 既存の文字列ベース evaluator は互換 fallback として維持しつつ、AST evaluator 側へ段階移行する
  - 既存 resolver は現時点では安全装置として必要であり、短期的には削除しない
  - 中長期的には、実データ観測を踏まえて担当範囲を縮小し、後方互換 fallback へ寄せる
  - 次候補:
    - `XLOOKUP` の binary search `search_mode=2/-2` の境界条件を必要に応じて詰める
    - 実データ頻出関数を AST evaluator 側へさらに寄せる
- merge multiline 次段タスク
  - `tests/fixtures/merge/merge-multiline-sample01.xlsx` は追加済み
  - 結合セル内改行の Markdown 正規化ポリシー変更を行う場合は、`markdown-normalize.ts` と `sheet-markdown.ts` と回帰テストをセットで見直す
- `セクション分割ブロック` の実装導入順を決める
- UI の `formulaDiagnostics` / `tableScores` 表示を見直す
- 表検出モードを追加する
  - 罫線を主手掛かりに表を検出するモードが欲しい
  - 非罫線ベースの検知では誤検知がつらいケースがある
  - 少なくとも `border` のような明示モードは検討したい
  - CLI / GUI の両方で切り替えられる形が望ましい
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
- Markdown 保存時の UTF-8 BOM 方針を整理する
  - 対象経路: GUI の `Save Markdown`、CLI の `--out`、ZIP 内 Markdown
  - Windows 既定アプリでの文字化け回避と、BOM を嫌う利用者・ツールの両方を考慮する
  - 既定値を BOM ありにするか、オプション化するかを決める
  - 3 経路で挙動をそろえるか、ZIP のみ別扱いにするかを決める
  - 仕様を決めたら README / spec / impl-spec / 回帰テストをまとめて更新する
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

- `local-data` の実データ差分レビュー継続
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
- `core.ts` に no-op の `extractSectionBlocks(...)` は追加済み
- `sheet-markdown.ts` の責務整理を実施し、セル整形 / narrative / section grouping / asset section / render state / summary の分割を進めた
- `tests/xlsx2md-sheet-markdown.test.js` に回帰テストを追加し、`sheet-markdown` 単体テストは 23 件通過を確認済み

## 参照

- 正本: `docs/TODO.md`
- 関連仕様:
  - `docs/xlsx2md-spec.md`
  - `docs/xlsx2md-impl-spec.md`
  - `docs/xlsx-formula-subset.md`
