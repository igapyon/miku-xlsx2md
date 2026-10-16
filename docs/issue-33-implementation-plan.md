# Issue #33 実装計画

作成日: 2026-10-16。状態: 実装完了。
実装担当: Codex（ユーザーの実装指示に基づく）。
対象: <https://github.com/igapyon/miku-xlsx2md/issues/33>

調査した HEAD は d77a4de9c2ca256a777ad8eef68cfbf02c3729bf。
Issue 本文の検証時 HEAD と一致する。計画作成時にはコード変更、fixture 作成、ビルド、テスト実行を行っていない。
実装開始時にブランチ `devel-tiga0702vec` と差分を確認した。適用される AGENTS.md は見つからなかった。

## 1. 採用する動作

基本の 2 ケースを回帰テストで固定する。追加の挙動変更は、balanced モードで、既に採用された表の直上にある「部分的な横結合を持ち、列を覆う見出し」を 1 行だけ含める処理とする。

| ID | 入力 | balanced の期待範囲 | 期待する内容 |
| --- | --- | --- | --- |
| W1 | 見出し A/B/C、データ 1/2/3/4、全体が無罫線 | A1:D2 | D1 は空見出し、D2 の 4 は表内 |
| W2 | W1 の存在する全セルに罫線 | A1:D2 | W1 と同じ |
| W3 | W1 の A1:C2 に罫線、D2 は無罫線 | A1:C2 | D2 の 4 は表外の本文に保持 |
| M1 | A1:B1 を結合して Group、C1 は C、下に数値 2 行、全体が無罫線 | A1:C3 | Group / [←M←] / C を見出しにする |
| M2 | M1 の A1:C3 全体に罫線 | A1:C3 | M1 と同じ |
| M3 | M1 の見出しは無罫線、A2:C3 に罫線 | A1:C3 | 今回改善するケース。最初の数値行をデータとして保持 |

W3 は単独の無罫線セルを隣接メモから区別する根拠が不足するため、現状の保守的な境界を仕様化する。値そのものの欠落と重複がないことを確認する。

改善対象は balanced のみ。M3 の border / planner-aware は現状どおり A2:C3 とする。
部分横結合のない通常の無罫線見出しや、複数段の見出し一般化は、この Issue の改善対象に含めない。
同じ配置の短い注記と見出しを意味で判別する機能も追加しない。本計画の条件を満たす配置は、文字列の意味にかかわらず見出しとして扱う。

## 2. 現行処理と変更位置

主な調査箇所:

- src/ts/table-detector.ts: detectTableCandidates / pruneRedundantCandidates / trimTableCandidateBounds / matrixFromCandidate
- src/ts/border-grid.ts: collectTableEdgeStats の rawBorderCount と borderCount
- src/ts/sheet-markdown.ts: collectSheetRenderState / createTableSections
- src/ts/markdown-export.ts: renderMarkdownTable
- tests/xlsx2md-table-detector.test.js: タイトル除外、注記除外、隣接表、各検出モードの既存テスト
- tests/xlsx2md-main.test.js: XLSX 解析から Markdown 出力までの既存テスト

現行では、罫線候補が広い候補の 40% 以上を占めると fallback を抑止する。また、先頭行が無罫線で値が 2 個以下なら trimTableCandidateBounds がその行を除外する。
M3 はこの両方の影響を受ける。

実装は、detectTableCandidates の既存の候補生成・trim・prune・モード固有の除外がすべて終了した後、最後の並び替えの前に追加する。
これにより、追加した見出しを再度 trim / prune で消す問題を避けられる。

既存の 40% / 60% 閾値、スコア重み、trimTableCandidateBounds、pruneRedundantCandidates の条件はそのまま利用する。
matrixFromCandidate と既存の merge token 描画を利用する。新しい CLI オプションや公開 API は追加しない。

## 3. 見出しを追加するための判定契約

以下をすべて満たすときだけ、候補の startRow を 1 減らす。
判定の入力は sheet、既に作成した cellMap、採用候補、採用候補全体の変更前の配列とする。
行・列は既存モデルと同じ 1 始まり。merge は数値の startRow / startCol / endRow / endCol を使う。

### 3.1 表本体

1. tableDetectionMode が balanced。
2. 本体が 2 行以上、3 列以上、startRow が 2 以上。
3. headerRow = candidate.startRow - 1 とする。これより上を探索しない。
4. 本体先頭行の rawBorderCount と nonEmptyCount が、いずれも Math.max(2, Math.ceil(width * 0.5)) 以上。

上記の統計は collectTableEdgeStats(cellMap, candidate.startRow, startCol, endCol) を使う。
既に採用済みの表に対する補正なので、見出しを加える前に 2 行の本体が必要である。
無罫線の結合見出しとデータ 1 行だけから表を新規成立させる拡張は行わない。

### 3.2 見出し行の罫線と幅

1. headerRow の startCol..endCol の rawBorderCount が 0。
2. 左右の直隣セルに値または自身の罫線があれば不採用。startCol が 1 の場合は左隣の確認を省略。
3. 候補を上へ 1 行広げた矩形が、他の採用候補と交差すれば不採用。矩形の交差は行・列の両区間が重なること。辺を共有する隣接セルは交差としない。

headerRow の borderCount は下の行の top 罫線を反映しうるので、「見出し自身に罫線がない」の判定に使わない。rawBorderCount を使う。
他候補との比較には全候補の変更前の座標を使う。配列の処理順で結果を変えない。

### 3.3 結合範囲

headerRow と候補の列範囲に交差する merge を列挙する。

1. 交差するすべての merge が、headerRow だけに収まり、候補の列範囲内に完全に収まること。
2. merge 同士が重ならないこと。
3. 少なくとも 1 個、幅が 2 以上かつ表全幅未満の横結合があること。
4. merge の左上以外に非空値がある場合は不採用。出力時に別の値を merge token で上書きする入力を取り込まない。

これにより、表全幅を結合した単独タイトル、縦結合、表幅を越える結合を取り込まない。
merge の ref 文字列を再解析する処理は追加しない。

### 3.4 見出しのラベルと列の対応

独立ラベルは、横結合の左上セル、または結合に含まれない値ありセルと定義する。
各ラベルについて次を要求する。

- String(cell.outputValue || "").trim() の長さが 1..24。
- 元の outputValue に CR または LF が含まれない。
- 現行のヘッダ評価と同じ正規表現 /^\d+(?:\.\d+)?$/ に一致しない。

結合左上のラベルはその結合列すべてを覆い、非結合のラベルは自分の 1 列を覆うものとして coverage を計算する。
独立ラベルが 2 個以上あり、startCol..endCol の全列が隙間なく覆われる場合だけ採用する。
欠けたセルや空文字は、結合に覆われていなければ coverage を満たさない。
[←M←] を判定用に作る必要はない。列の被覆は merge 情報から求める。

例: A1:B1 = Group、C1 = C は 2 ラベル・3 列被覆で採用。
例: A1:C1 = タイトル は 1 ラベルで不採用。
例: A1 = タイトル、B1 = 更新日、C1 = 日付 は、部分横結合がないため不採用。

## 4. コード変更の手順

変更の中心は src/ts/table-detector.ts の private helper と、detectTableCandidates の末尾とする。

1. private helper findBorderlessMergedHeaderRow を追加する。
   - 引数: sheet、cellMap、candidate、otherCandidates。
   - 返り値: 第 3 節の条件を満たす headerRow、または null。
   - 入力のセル、merge、候補、配列を変更しない。
   - モード判定は呼び出し側で行う。
2. 既存 finalCandidates を変更前の配列として保持し、balanced のときだけ map する。
3. 採用時は新しい候補オブジェクトと reasonSummary 配列を作り、startRow だけを更新する。
4. reasonSummary に次を 1 回だけ追加する。
   - Included borderless merged header row (no score change)
5. score は元の表本体の採用スコアを維持する。加点や再スコアリングはしない。
   - 現行も trim 後の矩形でスコアを再計算していない。
   - 既存候補が成立した根拠と、見出し追加の根拠を reasonSummary で説明する。
6. 最終範囲が同じ候補を完全一致の座標キーで重複除去する。重複時は先の候補を保持する。
7. 既存と同じ startRow / startCol 順で sort して返す。

処理の形:

~~~text
existing candidate generation
existing trim and candidate selection
existing prune and mode-specific filtering
originalCandidates = finalCandidates
if balanced:
    for each candidate in originalCandidates:
        row = findBorderlessMergedHeaderRow(... originalCandidates except self)
        if row exists:
            copy candidate with startRow = row and append reason
        else:
            keep candidate
deduplicate exact final bounds
sort and return
~~~

見出し追加後に trimTableCandidateBounds や pruneRedundantCandidates を再実行しない。
headerRow を再帰的・反復的に追加しない。
helper をテスト専用に module registry へ公開する必要はない。既存の detectTableCandidates 経由で振る舞いをテストする。

## 5. 最小 XLSX fixture

tests/fixtures/table/ に次の 6 ファイルを追加する。各ファイルは 1 シート、シート名 Sheet1。
テスト用のタイトル行は加えず、Issue と同じ A1 起点を維持する。目的はファイル名と tests/fixtures/README.md に記録する。

| ID | ファイル名 |
| --- | --- |
| W1 | table-header-width-unbordered-sample01.xlsx |
| W2 | table-header-width-bordered-sample01.xlsx |
| W3 | table-header-width-mixed-borders-sample01.xlsx |
| M1 | table-merged-header-unbordered-sample01.xlsx |
| M2 | table-merged-header-bordered-sample01.xlsx |
| M3 | table-merged-header-mixed-borders-sample01.xlsx |

W 系のセル: A1=A、B1=B、C1=C、A2=1、B2=2、C2=3、D2=4。D1 の c 要素は作らない。
M 系のセル: A1=Group、C1=C、A2=1、B2=2、C2=3、A3=4、B3=5、C3=6。mergeCell ref=A1:B1。

罫線は全辺 thin の 1 種類とし、各 ID の罫線配置を第 1 節に合わせる。
M2 は空の B1 も罫線付きセルとして作る。M1 / M3 の B1 は c 要素を省略し、merge 情報による被覆と token 展開を検証する。
文字列は inlineStr、数値は数値セルとする。数式や sharedStrings は不要。

fixture 作成は実装開始後に行う。生成方法は Python 標準ライブラリ zipfile でよい。
生成用スクリプトは workplace/issue-33/ 配下の作業用ファイルとし、実行環境の追加依存は導入しない。
ZIP の日時を固定する場合は 1980-01-01 00:00:00、ZIP_STORED を利用できる。
最小パーツは次の 6 個。

- [Content_Types].xml
- _rels/.rels
- xl/workbook.xml
- xl/_rels/workbook.xml.rels
- xl/styles.xml
- xl/worksheets/sheet1.xml

styles.xml には既定の font / fill / border / xf と、全辺 thin の border / xf を用意し、セルの s 属性で切り替える。
正しい Content-Type、workbook から worksheet と styles への relationships を設定する。
docProps、絶対パス、作者、アプリ情報、外部リンクを埋め込まない。

## 6. テスト設計

### 6.1 検出器のユニットテスト

tests/xlsx2md-table-detector.test.js の既存 createCell / buildCellMap / bootTableDetector を使う。
以下を detectTableCandidates 経由で追加する。期待範囲と候補数を明示し、該当ケースに意図しない候補が増えていないことも確認する。

| ケース | 期待 |
| --- | --- |
| 第 1 節の W1..M3 | 第 1 節の範囲 |
| 幅 4、A1:B1=Group1、C1:D1=Group2、下に罫線付き数値 2 行 | A1:D3 |
| M3 を B5:D7 に平行移動 | B5:D7 |
| A1:C1 は全幅タイトル、A2:B2=Group、C2=C、A3:C4 は罫線付き本体 | A2:C4。タイトルまで追加しない |
| A1:C1 は全幅の単独タイトル、A2:C3 は罫線付き本体 | A2:C3 |
| 幅 2 の本体の上に非結合のタイトルと日付 | 本体の範囲を維持 |
| 幅 3 の本体の上に非結合の短文 3 個 | 本体の範囲を維持 |
| 幅 4 で A1:B1=Group、C1 空欄、D1=D | A2:D3。被覆の穴で不採用 |
| M3 の Group を X を 25 個並べた文字列、Group + LF + Name、123 にそれぞれ変更 | A2:C3 |
| M3 の結合左上 A1 を空欄に変更 | A2:C3 |
| M3 の結合従属セル B1 に値を追加 | A2:C3 |
| M3 の結合を A1:D1 に変更 | A2:C3。表幅を越える結合を不採用 |
| M3 の結合を A1:B2 に変更 | A2:C3。縦方向に交差する結合を不採用 |
| M3 に D1=Extra を追加 | A2:C3。見出しの連続した張り出しで不採用 |
| M3 を A1:C3 と E1:G3 に配置し、D 列を空にする | A1:C3 と E1:G3 の 2 表を維持 |

ラベル検証の上限は、24 文字なら採用、25 文字なら不採用の両方を確認する。
新しい reason は M3 の balanced で 1 個だけ付与されること。
score は、同じ本体から見出し行だけ除いた入力の候補スコアと等しいこと。スコアの数値そのものを固定しない。
入力を変更しないことは、代表ケースで入力の前後比較を行う。

既存のタイトル除外、後続注記除外、隣接表、疎な結合フォーム、内側の表、planner-aware のテストを維持する。

### 6.2 XLSX から Markdown までの回帰テスト

新規 tests/xlsx2md-table-header-regression.test.js を作る。
tests/xlsx2md-main.test.js と同じ jsdom 環境、bootXlsx2mdCore、ArrayBuffer 変換を使い、fixture 6 個を実際に parseWorkbook で読む。
変換オプションは balanced、treatFirstRowAsHeader=true、trimText=true、removeEmptyRows=true、removeEmptyColumns=true を明示する。

各ケースで以下を確認する。

- パースされた元セル値、M 系の merge ref=A1:B1。
- summary.tables が 1。
- summary.tableScores の range。表示形式は A1-D2 / A1-C3 / A1-C2 のようにハイフンになる。
- 表の行列の完全一致。Markdown の | で始まる行だけを取り出すと、front matter や見出しの差分を避けられる。
- W3 は表の外の行に単独の 4 が残り、表内に 4 が追加されていないこと。
- M3 は Group が本文全体で 1 回だけ現れ、数値 1/2/3 と 4/5/6 がともにデータ行にあること。

W1 / W2 の表行:

~~~markdown
| A | B | C |  |
| --- | --- | --- | --- |
| 1 | 2 | 3 | 4 |
~~~

W3 の表行（表外の単独の 4 と合わせて確認）:

~~~markdown
| A | B | C |
| --- | --- | --- |
| 1 | 2 | 3 |
~~~

M1 / M2 / M3 の表行:

~~~markdown
| Group | [←M←] | C |
| --- | --- | --- |
| 1 | 2 | 3 |
| 4 | 5 | 6 |
~~~

モード差分を次の表どおり追加確認する。

| ID | balanced | border | planner-aware |
| --- | --- | --- | --- |
| W1 | A1:D2 | 表なし | A1:D2 |
| W2 | A1:D2 | A1:D2 | A1:D2 |
| W3 | A1:C2 | A1:C2 | A1:C2 |
| M1 | A1:C3 | 表なし | A1:C3 |
| M2 | A1:C3 | A1:C3 | A1:C3 |
| M3 | A1:C3 | A2:C3 | A2:C3 |

M3 の balanced について treatFirstRowAsHeader=false も確認する。
その場合は空の Markdown 見出しに続いて Group / [←M←] / C の行、数値の 2 行がすべてデータとして現れること。
removeEmptyRows=false / removeEmptyColumns=false の組み合わせでも、M3 の表範囲と 3 列が保たれることを代表 1 ケースで確認する。

### 6.3 実装後の検証手順

実装時に次の順で実行した。

~~~sh
npm test -- tests/xlsx2md-table-detector.test.js tests/xlsx2md-table-header-regression.test.js tests/xlsx2md-main.test.js tests/xlsx2md-fixture-hygiene.test.js
npm test
~~~

npm test の pretest が build:core を実行するので、dist/js の古い成果物を直接テストしない。
局所テストが通ったら全体テストを 1 回行う。追加変更や失敗があった場合に必要な範囲を再実行する。
fixture hygiene の既存 3 テストも合格させる。

## 7. 変更ファイルと実施内容

予定する変更:

- src/ts/table-detector.ts
- tests/xlsx2md-table-detector.test.js
- tests/xlsx2md-table-header-regression.test.js（新規）
- tests/fixtures/table/ の第 5 節の 6 ファイル（新規）
- tests/fixtures/README.md: fixture の目的、条件、期待範囲
- docs/xlsx2md-spec.md: 7.4 付近に balanced の限定的な見出し追加条件と W3 の境界方針
- docs/xlsx2md-impl-spec.md: 9 章に、候補選択後の 1 行追加とスコア維持の説明

実施内容:

1. HEAD、作業ツリー、対象ソース、適用される AGENTS.md を確認した。
2. 6 個の XLSX fixture と、実解析から Markdown までの回帰テストを追加した。
3. 検出器の正例・負例テストを追加した。
4. 第 3・4 節に沿って private helper と候補後処理を実装した。
5. モード差、オプション差、表外値、結合マーカー、スコア維持を確認した。
6. fixture README と仕様 2 文書を更新した。
7. 対象テストと全体テストを実行し、差分を確認した。

別のケースの失敗を直すために閾値や既存の期待値を一括変更しない。
既存ケースと衝突したら、その入力を確認して局所条件で解決できるかを報告する。
生成される dist/js は .gitignore 対象であり、正本は src/ts とする。

## 8. 完了条件

- [x] 基本 4 ケースと罫線混在 2 ケースの実ファイル回帰テストがある。
- [x] M3 の balanced が A1:C3 となり、結合マーカーと 2 データ行が保持される。
- [x] W3 の D2 は表外に 1 回保持される。
- [x] 単独タイトル、部分被覆、幅超過、縦結合、隣接表について期待した境界になる。
- [x] border / planner-aware の上表の期待範囲を維持する。
- [x] 全体テストと fixture hygiene が通る。
- [x] fixture README と仕様に実装後の条件と制約が記載されている。
- [x] 実装報告に実行したテストと結果を記載する。

検証記録:

- `npm test -- tests/xlsx2md-table-detector.test.js tests/xlsx2md-table-header-regression.test.js tests/xlsx2md-main.test.js tests/xlsx2md-fixture-hygiene.test.js` — 4 files / 88 tests passed.
- `npm test` — 37 files / 351 tests passed.

GitHub Issue の本文更新、コメント、クローズ、コミット、push は、各操作の依頼があった段階で miku-scm の対応するワークフローへ進む。
