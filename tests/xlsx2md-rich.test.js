// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootXlsx2mdCore } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootCore() {
  return bootXlsx2mdCore(__dirname);
}

function toArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function loadFixtureMarkdown(fixtureName, formattingMode) {
  const api = bootCore();
  const fixturePath = path.resolve(__dirname, `./fixtures/rich/${fixtureName}`);
  const fixtureBytes = readFileSync(fixturePath);
  const workbook = await api.parseWorkbook(toArrayBuffer(fixtureBytes), fixtureName);
  const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
    treatFirstRowAsHeader: true,
    trimText: true,
    removeEmptyRows: true,
    removeEmptyColumns: true,
    formattingMode
  })[0];
  return { workbook, markdownFile };
}

describe("xlsx2md rich fixtures", () => {
  it("renders rich-text-github-sample01.xlsx in github mode with inline styling and <br>", async () => {
    const { workbook, markdownFile } = await loadFixtureMarkdown("rich-text-github-sample01.xlsx", "github");

    expect(workbook.sheets).toHaveLength(1);
    expect(markdownFile.summary.formattingMode).toBe("github");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("**bold whole cell**");
    expect(markdownFile.markdown).toContain("*italic whole cell*");
    expect(markdownFile.markdown).toContain("~~strike whole cell~~");
    expect(markdownFile.markdown).toContain("<ins>underline whole cell</ins>");
    expect(markdownFile.markdown).toContain("plain **bold** *italic* strike <ins>underline</ins>");
    expect(markdownFile.markdown).toContain("***bold+italic***");
    expect(markdownFile.markdown).toContain("**<ins>bold+underline</ins>**");
    expect(markdownFile.markdown).toContain("*~~italic+strike~~*");
    expect(markdownFile.markdown).toContain("改行入り文字列で<br>**一部だけ太**字");
    expect(markdownFile.markdown).toContain("重要, <ins>下線</ins>,~~取消線~~,**強調**");
    expect(markdownFile.markdown).toContain("**12345**");
    expect(markdownFile.markdown).toContain("<ins>24690</ins>");
  });

  it("renders rich-text-github-sample01.xlsx in plain mode without inline styling or <br>", async () => {
    const { markdownFile } = await loadFixtureMarkdown("rich-text-github-sample01.xlsx", "plain");

    expect(markdownFile.summary.formattingMode).toBe("plain");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("bold whole cell");
    expect(markdownFile.markdown).toContain("italic whole cell");
    expect(markdownFile.markdown).toContain("strike whole cell");
    expect(markdownFile.markdown).toContain("underline whole cell");
    expect(markdownFile.markdown).toContain("改行入り文字列で 一部だけ太字");
    expect(markdownFile.markdown).not.toContain("<ins>underline whole cell</ins>");
    expect(markdownFile.markdown).not.toContain("**bold whole cell**");
    expect(markdownFile.markdown).not.toContain("<br>");
  });

  it("renders rich-markdown-escape-sample01.xlsx in github mode and keeps fixture-specific cases stable", async () => {
    const { workbook, markdownFile } = await loadFixtureMarkdown("rich-markdown-escape-sample01.xlsx", "github");

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].cells.length).toBe(36);
    expect(markdownFile.summary.formattingMode).toBe("github");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("line1 \\* x<br>**line2 \\[y\\]\\(z\\)**");
    expect(markdownFile.markdown).toContain("| Header \\\\| One | Header \\*Two\\* | Header \\[Three\\]\\(x\\) |");
    expect(markdownFile.markdown).toContain("| a**\\*b** | a\\_**b** | a\\~\\~b\\~\\~c |");
    expect(markdownFile.markdown).toContain("| \\# not **heading** | \\- not list | 1\\. ***not*** list |");
    expect(markdownFile.markdown).toContain("| a\\*b | **a\\*b** |");
    expect(markdownFile.markdown).toContain("| a\\_b | *a\\_b* |");
    expect(markdownFile.markdown).toContain("| a\\~\\~b\\~\\~c | ~~a\\~\\~b\\~\\~c~~ |");
    expect(markdownFile.markdown).toContain("| \\# not heading | <ins>\\# not heading</ins> |");
    expect(markdownFile.markdown).toContain("| &lt;tag&gt; | &lt;tag&gt; |");
    expect(markdownFile.markdown).toContain("| \\!\\[alt\\]\\(image.png\\) | \\!\\[alt\\]\\(image.png\\) |");
    expect(markdownFile.markdown).toContain("| code \\`sample\\` | code \\`sample\\` |");
    expect(markdownFile.markdown).toContain("| path\\\\to\\\\file | path\\\\to\\\\file |");
  });

  it("renders rich-markdown-escape-sample01.xlsx in plain mode as plain text without <br>", async () => {
    const { markdownFile } = await loadFixtureMarkdown("rich-markdown-escape-sample01.xlsx", "plain");

    expect(markdownFile.summary.formattingMode).toBe("plain");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("line1 \\* x line2 \\[y\\]\\(z\\)");
    expect(markdownFile.markdown).not.toContain("<br>");
    expect(markdownFile.markdown).toContain("| a\\*b | a\\*b |");
    expect(markdownFile.markdown).toContain("| a\\_b | a\\_b |");
    expect(markdownFile.markdown).toContain("| a\\~\\~b\\~\\~c | a\\~\\~b\\~\\~c |");
    expect(markdownFile.markdown).toContain("| \\# not heading | \\# not heading |");
    expect(markdownFile.markdown).toContain("| a\\*b | a\\_b | a\\~\\~b\\~\\~c |");
    expect(markdownFile.markdown).toContain("| \\# not heading | \\- not list | 1\\. not list |");
    expect(markdownFile.markdown).toContain("| \\!\\[alt\\]\\(image.png\\) | \\!\\[alt\\]\\(image.png\\) |");
    expect(markdownFile.markdown).toContain("| code \\`sample\\` | code \\`sample\\` |");
    expect(markdownFile.markdown).toContain("| path\\\\to\\\\file | path\\\\to\\\\file |");
  });

  it("renders rich-usecase-sample01.xlsx in github mode with hyperlinks, inline styling, and <br>", async () => {
    const { workbook, markdownFile } = await loadFixtureMarkdown("rich-usecase-sample01.xlsx", "github");

    expect(workbook.sheets).toHaveLength(1);
    expect(markdownFile.summary.formattingMode).toBe("github");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.markdown).toContain("| [Apple](https://www.apple.com/) | ***Apple*** の製品が<ins>購入できます</ins>。 |");
    expect(markdownFile.markdown).toContain("| [Google](https://www.google.com/) | とても<ins>有名</ins>な**検索サイト**です。 |");
    expect(markdownFile.markdown).toContain("| [Amazon](https://www.amazon.co.jp/) | **<ins>お買い物</ins>**でお世話になっています。 |");
    expect(markdownFile.markdown).toContain("| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに<br>**ネットショップ**でもお世話になっています。 |");
    expect(markdownFile.markdown).toContain("~~池袋の激戦区で、生き残るのはどの店舗か。~~<br>→トルツメ: この部分は文面から外すことを提案。");
  });

  it("renders rich-usecase-sample01.xlsx in plain mode with links preserved and styling removed", async () => {
    const { markdownFile } = await loadFixtureMarkdown("rich-usecase-sample01.xlsx", "plain");

    expect(markdownFile.summary.formattingMode).toBe("plain");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.markdown).toContain("| [Apple](https://www.apple.com/) | Apple の製品が購入できます。 |");
    expect(markdownFile.markdown).toContain("| [Google](https://www.google.com/) | とても有名な検索サイトです。 |");
    expect(markdownFile.markdown).toContain("| [Amazon](https://www.amazon.co.jp/) | お買い物でお世話になっています。 |");
    expect(markdownFile.markdown).toContain("| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに ネットショップでもお世話になっています。 |");
    expect(markdownFile.markdown).toContain("池袋の激戦区で、生き残るのはどの店舗か。 →トルツメ: この部分は文面から外すことを提案。");
    expect(markdownFile.markdown).not.toContain("<br>");
    expect(markdownFile.markdown).not.toContain("<ins>");
    expect(markdownFile.markdown).not.toContain("~~池袋の激戦区で、生き残るのはどの店舗か。~~");
  });
});
