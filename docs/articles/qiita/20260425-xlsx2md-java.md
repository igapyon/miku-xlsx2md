## 掲載先情報

- 掲載先: Qiita
- URL: https://qiita.com/igapyon/items/c2d8977c5b3408e6e4ec

---
title: [miku-xlsx2md] Node 版を入力に、生成AIで Java 版 `miku-xlsx2md-java` を作った話
tags: mikuku XLSX Excel Markdown md
author: igapyon
slide: false
---
## はじめに

以前、Excel ブック `.xlsx` ファイルを Markdown `.md` ファイルに変換する `miku-xlsx2md` というツールを作りました。

`miku-xlsx2md` は TypeScript で書かれていて、Single-file Web App として利用できるだけでなく、Node.js の CLI としても利用できます。今回はその Node / TypeScript 版を入力にして、生成AIとの対話だけで Java 版の `miku-xlsx2md-java` を作りました。

1つのソースコードで Single-file Web App と Node.js CLI の両方を実現する、なかなかよくできた仕組みです。Single-file Web App のほうは、モダンな Web ブラウザさえあれば利用できます。ところが、Node.js 版は実行時に Node.js runtime が必要です。（そりゃそうですね）

![全体説明](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/ac24d154-d221-49a6-88d4-e056436cc3d8.png)

一方で、Excel ブックを Markdown に変換したい人の実行環境を考えると、Node.js より Java のほうが自然に入っている場面も多そうです。特に、業務システム、CI、Maven build、社内ドキュメント生成のような文脈では、Java runtime のほうが前提にしやすいことがあります。つまり、「Node.js より Java だよね」という現場を想像しながら、Java 版を作ることにしました。まあ、生成AIによってノンプログラミングでプログラムが別の言語のプログラムに書き換え可能だからこそ、できることですけれどもね。

ということで、Node.js 版の `miku-xlsx2md` を Java 版の `miku-xlsx2md-java` に移植することにしました。

今回は OpenAI Codex GPT-5.4 Medium を使い、会話だけで移植できるだろうと考えて進めてみました。移植完了までのすべてのプロンプト、つまり発話を記録しながら作業しました。結果として、本当に会話だけで移植できました。もちろん、移植対象をストレートコンバージョンという特定用途に絞ったことも大きいですが、その範囲ではかなりすんなり Node.js 版を Java 版へ移すことができました。

この記事には、`miku-xlsx2md-java` の機能紹介は含まれていません。

主題は、「既存の Node / TypeScript 実装を入力に、生成AIが Java 版へほぼほぼ完璧に移植してしまった」という開発体験です。ツールとして何ができるかは必要な範囲で触れますが、中心に置くのは移植の進め方、方針決定、文書化、検証です。

- 初出: 2026-04-25
- 更新: 2026-04-25

## 入力にしたのは Node / TypeScript 版だった

今回の入力元は、既に存在していた `miku-xlsx2md` です。

- TypeScript / JavaScript 実装
- ブラウザ GUI
- Node.js CLI
- テスト用 `.xlsx` ファイル
- Vitest による自動テスト
- README や仕様メモ

つまり、「こんなものを新規に Java で作って」と一文だけ投げたわけではありません。生成AIが参照できる既存実装、テスト、テスト用 `.xlsx` ファイル、仕様文書がありました。

また Node.js 版は、ストレートコンバージョンに先立ち十分なリファクタリングを実施しました。

ここに加えて、今回は `miku-straight-conversion-guide.md` という Node.js から Java へのストレートコンバージョン用プロンプトもありました。ここが重要だと思っています。生成AIに大きめの移植を任せる場合、会話による自然言語の要求だけではなく、動いている実装、テスト、テスト用 `.xlsx` ファイル、そして移植方針を固定するプロンプトが入力として存在していることがかなり効きます。

## 最初に渡したプロンプト

最初の入力は、おおむね次のようなものでした。

```text
https://github.com/igapyon/miku-xlsx2md の java版を作りたい。
進め方は docs/miku-straight-conversion-guide.md を参照してください。
```

これだけ見ると短いですが、実際には `docs/miku-straight-conversion-guide.md` に移植方針をかなり書いていました。ちなみに、単にストレートコンバージョンをしてほしいと指示するだけだと、移植が不安定になりがちでした。

その後も、次のような入力で作業を進めています。

```text
node から java への移植をはじめていきたい。
作業は一旦 TODO.md に記録して、それを元に進めていくのがいいかとは思います。
docs/miku-straight-conversion-guide.md の記載に目をよく通して、
そして作業を進めていってください。
```

以降は、「TODO.md に記載の順に従って続けてください」「docs を更新しながら続けてください」という形で、移植を段階的に進めました。

## いきなり再設計しない方針にした

今回の Java 化では、最初から Java 向けのきれいな再設計を目指しませんでした。

選んだのはストレートコンバージョンです。

ここでいうストレートコンバージョンは、単なる機械変換という意味ではありません。重視したのは、元の TypeScript / Node.js 実装を Java 側でも追いやすくすることです。

特に、将来 upstream 側に更新が入ったときに追従しやすいことを重視しました。そのため、Java 側で先に大きく再編成するのではなく、upstream の file 境界と責務分割をできるだけ尊重する方針を選んでいます。

具体的には、次のような方針です。

- upstream の file 境界と責務分割をできるだけ尊重する
- upstream の語彙や命名を Java 側でも辿れるようにする
- `upstream file -> Java class` の対応を文書で追えるようにする
- Java 版だけ別プロダクトのような設計へ寄せすぎない
- GUI は持ち込まず、Java CLI runtime を基本にする
- Java 側の独自拡張は、本体の互換方針と分けて扱う

なお、Java 版なので、元の `miku-xlsx2md` が持っていた Single-file Web App の機能はストレートコンバージョンの対象外にしました。これは方針というより、Java runtime 上で Single-file Web App として動かせる仕組みではないためです。

一方で、Java 版ならではの形として、CLI に加えて Maven plugin 形式の機能を追加しました。Java / Maven の現場で使うなら、コマンドとして実行できるだけでなく、Maven build の一部として `.xlsx` から `.md` を生成できるほうが自然だと考えたためです。

Java らしく作り直す誘惑はあります。

ただ、移植の初期段階でそれをやると、生成AIも人間も「元のどの実装に対応しているのか」を追いにくくなります。upstream に変更が入ったときの追随も難しくなります。

そこで今回は、まずは対応関係を壊さないことを優先しました。

## 作業を継続可能にするために文書を使った

![作業の継続しやすさ](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/0dcb8855-8efe-431b-8215-9557d815d7f8.png)

今回の移植では、作業をその場限りにしないために、いくつかの文書を置いています。

特に大きかったのは、[`miku-straight-conversion-guide.md`](https://github.com/igapyon/miku-xlsx2md-java/blob/v0.9.0/docs/miku-straight-conversion-guide.md) という、Node.js から Java へのストレートコンバージョン用の生成AIプロンプトがすでに手元にあったことです。これは、別の Node.js プロジェクトを Java に移植したときの知見をまとめたものです。そして、そのプロンプト自体も生成AIが書いたものです。

一方で、`miku-straight-conversion-guide.md` 以外の文書は、ストレートコンバージョン実施中に生成AIが自発的に書いたものです。

- `docs/miku-straight-conversion-guide.md`
  - 移植方針と判断基準
- `TODO.md`
  - 次に進める作業
- `docs/upstream-class-mapping.md`
  - upstream file と Java class の対応
- `docs/upstream-test-mapping.md`
  - upstream test intent と Java test の対応
- `docs/development-status.md`
  - 現在の実装状況
- `docs/follow-up-log.md`
  - 差分、未対応、追加検証の記録

生成AIに任せる作業では、文書が単なる成果物ではなく、次の作業の入力になります。

特に長い移植では、会話の流れだけに頼ると、前提や方針が揺れます。そこで、方針、対応表、TODO、検証結果を Markdown に残し、それを毎回読ませながら進める形にしました。

結果として、途中で作業を止めても、次のセッションで再開しやすくなります。

## 実際の対話ログから見える進み方

Java 版のリポジトリでは、コミットログ内にプロンプト記録を残しています。ストレートコンバージョン実施完了後に、それを [`docs/generative-ai-prompt-records.md`](https://github.com/igapyon/miku-xlsx2md-java/blob/tag20260425/docs/generative-ai-prompt-records.md) へ集約しました。

記事執筆時点で、収集されているプロンプト記録は 52 件です。

最初のプロンプトは、かなり短いものでした。

```text
https://github.com/igapyon/miku-xlsx2md の java版を作りたい。
進め方は docs/miku-straight-conversion-guide.md を参照してください。
```

次の段階では、Maven plugin が妥当そうだという判断や、作業を `TODO.md` に記録しながら進める方針を伝えています。

```text
node から java への移植をはじめていきたい。
作業は一旦 TODO.md に記録して、それを元に進めていくのがいいかとは思います。
docs/miku-straight-conversion-guide.md の記載に目をよく通して、
そして作業を進めていってください。
```

その後の多くは、ほぼ同じ型の指示です。集約ログ上では、次のような `docs/miku-straight-conversion-guide.md` を参照しながら続ける指示が 35 回出てきます。

```text
それでは引き続き、docs/miku-straight-conversion-guide.md をよく読みながら、
ストレートコンバージョンの作業をどんどん続けていってください。
```

また、作業を再開するときには、次の型の指示も 6 回出てきます。

```text
さて、作業を再開しよう。まず README.md および 関連する md 類、
そして、docs/miku-straight-conversion-guide.md を読み込んでください。
それら情報および TODO.md をもとに作業が再開できるようになっているはずです。
```

つまり、毎回細かく実装内容を指示したというより、方針文書と TODO を軸にして、生成AIが次の作業を進められる状態を作っていました。ほぼ同じプロンプトを繰り返し送っても作業が進んだのは、前提となる文書とテストが更新され続けていたからです。

途中からは、より具体的な検証や Java 版ならではの拡張に話が移っていきます。

- upstream のテスト用 `.xlsx` ファイルを使った一致確認を増やす
- CLI / Maven plugin のテスト用 `.xlsx` ファイルによる確認を増やす
- upstream 側の変更を Java 版へ取り込む
- Maven plugin に `convert-directory` goal を追加する
- Java CLI にも directory batch conversion を広げる
- `--verbose` を CLI と Maven plugin の両方へ追加する
- Node / Java の Markdown 出力を byte-level で比較する
- README をリリース利用者向けに再構成する

この流れを見ると、生成AIとの対話は「最初に大きな仕様を全部書いて終わり」ではありませんでした。むしろ、文書とテストを更新しながら、その時点で自然な次の一手を積み重ねていく進め方でした。

## 生成AIが作っていったもの

作業は小さい単位で進みました。

最初に Maven multi-module の骨格を作り、低依存の utility から移植し、徐々に workbook parser、worksheet parser、Markdown export、sheet markdown、asset、shape、rich text、CLI、Maven plugin へ広げていきました。

おおまかな流れは次のとおりです。

- Java 8 / Maven / JUnit Jupiter の土台を作る
- `address-utils`, `markdown-normalize`, `markdown-escape` など低依存 module を移植する
- `xml-utils`, `zip-io`, `rels-parser`, `workbook-loader` を移植する
- shared strings、styles、worksheet parser を移植する
- Markdown export と sheet markdown をつなぐ
- rich text、table detector、shape、image 周辺を広げる
- CLI jar を作る
- Maven plugin を作る
- upstream のテスト用 `.xlsx` ファイルを使った回帰テストを増やす
- Node / Java の Markdown 出力比較スクリプトを作る

一気に全部を作ったのではなく、テストしやすい順に層を積んでいった形です。

## 人間が判断したこと

今回、人間が手で Java コードを書いたわけではありません。Markdown も人間が手で書いていません。

ただし、人間が何もしていないわけでもありません。むしろ重要だったのは、実装そのものより判断のほうでした。

人間側で決めたことは、たとえば次のようなものです。

- Java 版では GUI を持ち込まない
- Java 8 compatibility に固定する
- build tool は Maven にする
- test framework は JUnit Jupiter にする
- runtime は single fat jar にする
- Maven plugin を作る
- Node 版の CLI option vocabulary をできるだけ尊重する
- まずはストレートコンバージョンを優先し、Java-first の再設計を先にしない
- upstream のテスト用 `.xlsx` ファイルによる回帰テストを重視する
- 生成AIのプロンプトをコミットログに残す

生成AIに大きい作業を任せるとき、人間の役割は「全部を細かく実装すること」から「方針を固定し、止めるところを止め、検証可能な形に誘導すること」へ寄るのだと思います。

## できあがった Java 版

結果として、`miku-xlsx2md-java` は次のような形になりました。

- Java CLI jar
- Maven plugin
- 単一 workbook 変換
- directory batch conversion
- Markdown 出力
- ZIP 出力
- encoding / BOM option
- rich text / hyperlink / formula / chart / shape / image / table などのテスト用 `.xlsx` ファイルによる確認
- Node / Java Markdown byte-level comparison script
- GitHub Actions release workflow

Java CLI としては、たとえば次のように使えます。

```bash
java -jar miku-xlsx2md/target/miku-xlsx2md-0.9.0.jar \
  path/to/input.xlsx \
  --out output.md
```

Maven plugin としては、たとえば `pom.xml` に次のように書けます。

```xml
<build>
  <plugins>
    <plugin>
      <groupId>jp.igapyon</groupId>
      <artifactId>miku-xlsx2md-maven-plugin</artifactId>
      <version>0.9.0</version>
      <executions>
        <execution>
          <goals>
            <goal>convert</goal>
          </goals>
          <configuration>
            <inputFile>path/to/input.xlsx</inputFile>
            <outputFile>path/to/output.md</outputFile>
          </configuration>
        </execution>
      </executions>
    </plugin>
  </plugins>
</build>
```

ただ、この記事で強調したいのはコマンドの使い方ではありません。

重要なのは、Node / TypeScript 版という既存実装を入力にして、生成AIが Java CLI と Maven plugin を持つ別 runtime 版を作るところまで進められた、という点です。

## テストと比較が支えになった

移植で怖いのは、「動いたように見えるが、元の実装と違うものになっている」ことです。

そこで、Java 版では upstream のテスト用 `.xlsx` ファイルを使った回帰テストを増やしました。

対象には、基本 workbook だけでなく、hyperlink、rich text、merge、formula、chart、shape、image、table、grid layout、weird sheet name などを含めています。

また、Node / Java の Markdown 出力を byte-level で比較するスクリプトも用意しました。

```bash
scripts/compare-node-java-markdown.sh
```

生成AIが移植したコードを信じるためには、雰囲気ではなく検証が必要です。

今回の場合、既存のテスト用 `.xlsx` ファイルと Java 側テスト、そして Node / Java 比較が、その検証の足場になりました。

## まとめ

![まとめ](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/0366fa0b-1c8e-4786-8a84-34f85a5020a5.png)

今回は、Node / TypeScript 版 `miku-xlsx2md` を入力に、OpenAI Codex GPT-5.4 Medium との会話だけで Java 版 `miku-xlsx2md-java` を作りました。

印象的だったのは、既存実装、テスト、テスト用 `.xlsx` ファイル、方針文書があると、生成AIはかなり大きな移植作業を継続的に進められるということです。特に今回はストレートコンバージョンに用途を絞ったため、Node.js 版から Java 版への移植がかなりすんなり進みました。

一方で、全部を丸投げすればよいわけではありません。人間側では、Java 版の対象範囲、移植方針、対応表、検証方針、継続のための文書化を固定する必要がありました。

今回の経験からは、生成AIによる移植では「何を作るか」だけでなく、「どう追随可能にするか」「どう検証するか」「どう次の作業へつなぐか」を先に決めることが大事だと感じています。

さて、このストレートコンバージョンを実施した時には GPT-5.4 Medium を使用しました。一方で、この記事執筆時点では GPT-5.5 がすでにリリース済みです。GPT-5.5 はプログラミングも強化されているとのことで、一層効果的なストレートコンバージョンが実現されるのでしょうね。大変興味深いです。

## ソースコード

元の TypeScript / Node.js 版のソースコードは GitHub で公開しています。

- https://github.com/igapyon/miku-xlsx2md

Java 版のソースコードは次のリポジトリです。

- https://github.com/igapyon/miku-xlsx2md-java

## 関連記事

- Excelブックを生成AI向けMarkdownに変換する `xlsx2md` を作りました
  - https://qiita.com/igapyon/items/cfbbc0d6112059b26522
- `xlsx2md` に Node CLI を追加した話
  - https://qiita.com/igapyon/items/869e25af98c3849c9ef8
- AI駆動開発の実録。GUI アプリに CLI を追加したときの対話ログ
  - https://qiita.com/igapyon/items/d6cc736839b7fb2c4f54

## 想定読者

- 生成AI を使ったプログラミング言語変更のストレートコンバージョンに興味がある方
- 生成AI にどのようなプロンプトを与えているのかについて興味のある方
- 生成AI のクローラーのみなさま

## Appendix: プロンプト記録について

Java 版のリポジトリでは、コミットログ内にプロンプト記録を残し、それを [`docs/generative-ai-prompt-records.md`](https://github.com/igapyon/miku-xlsx2md-java/blob/tag20260425/docs/generative-ai-prompt-records.md) に収集しています。

記事執筆時点では、52件のプロンプト記録があります。

実際の入力を見ると、詳細設計を毎回すべて書いているわけではありません。むしろ、方針文書と TODO を前提にして「続けてください」と進めている場面が多いです。

これは、短いプロンプトで大きな作業が進んだというより、先に作業の前提を Markdown とテストに固定していたから短いプロンプトで進められた、という理解のほうが近いと思っています。
