// @vitest-environment jsdom

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootRegisteredModule } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootNarrativeStructure() {
  return bootRegisteredModule(__dirname, [
    "src/js/markdown-normalize.js",
    "src/js/narrative-structure.js"
  ], "narrativeStructure");
}

describe("xlsx2md narrative structure", () => {
  it("renders an indented parent-child block as heading plus bullets", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 3,
      lines: ["通常のテキスト", "字下げされたテキスト", "テキスト"],
      items: [
        { row: 1, startCol: 1, text: "通常のテキスト", cellValues: ["通常のテキスト"] },
        { row: 2, startCol: 2, text: "字下げされたテキスト", cellValues: ["字下げされたテキスト"] },
        { row: 3, startCol: 2, text: "テキスト", cellValues: ["テキスト"] }
      ]
    });

    expect(markdown).toContain("### 通常のテキスト");
    expect(markdown).toContain("- 字下げされたテキスト");
    expect(markdown).toContain("- テキスト");
  });

  it("renders flat narrative rows as plain paragraphs", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      lines: ["一行目", "二行目"],
      items: [
        { row: 1, startCol: 1, text: "一行目", cellValues: ["一行目"] },
        { row: 2, startCol: 1, text: "二行目", cellValues: ["二行目"] }
      ]
    });

    expect(markdown).toBe("一行目\n\n二行目");
  });

  it("normalizes heading and list markers inside hierarchical narrative items", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 3,
      lines: ["### 親", "- 子", "1. 孫"],
      items: [
        { row: 1, startCol: 1, text: "### 親", cellValues: ["### 親"] },
        { row: 2, startCol: 2, text: "- 子", cellValues: ["- 子"] },
        { row: 3, startCol: 2, text: "1. 孫", cellValues: ["1. 孫"] }
      ]
    });

    expect(markdown).toContain("### 親");
    expect(markdown).toContain("- 子");
    expect(markdown).toContain("- 孫");
  });

  it("starts a new heading when indentation returns to the parent level", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 5,
      lines: ["親1", "子1", "子2", "親2", "子3"],
      items: [
        { row: 1, startCol: 1, text: "親1", cellValues: ["親1"] },
        { row: 2, startCol: 2, text: "子1", cellValues: ["子1"] },
        { row: 3, startCol: 2, text: "子2", cellValues: ["子2"] },
        { row: 4, startCol: 1, text: "親2", cellValues: ["親2"] },
        { row: 5, startCol: 3, text: "子3", cellValues: ["子3"] }
      ]
    });

    expect(markdown).toContain("### 親1");
    expect(markdown).toContain("- 子1");
    expect(markdown).toContain("- 子2");
    expect(markdown).toContain("### 親2");
    expect(markdown).toContain("- 子3");
  });

  it("detects a heading block only when the second item is indented deeper", () => {
    const api = bootNarrativeStructure();

    expect(api.isSectionHeadingNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      lines: ["親", "子"],
      items: [
        { row: 1, startCol: 1, text: "親", cellValues: ["親"] },
        { row: 2, startCol: 2, text: "子", cellValues: ["子"] }
      ]
    })).toBe(true);

    expect(api.isSectionHeadingNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      lines: ["一行目", "二行目"],
      items: [
        { row: 1, startCol: 1, text: "一行目", cellValues: ["一行目"] },
        { row: 2, startCol: 1, text: "二行目", cellValues: ["二行目"] }
      ]
    })).toBe(false);
  });

  it("renders calendar-like narrative rows with cell boundaries preserved", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 4,
      lines: ["2021年1月", "日 月 火 水 木 金 土", "2021-01-03 2021-01-04 2021-01-05 2021-01-06 2021-01-07 2021-01-08 2021-01-09", "仕事 私用 その他"],
      items: [
        { row: 1, startCol: 1, text: "2021年1月", cellValues: ["2021年1月"] },
        { row: 2, startCol: 1, text: "日 月 火 水 木 金 土", cellValues: ["日", "月", "火", "水", "木", "金", "土"] },
        { row: 3, startCol: 1, text: "2021-01-03 2021-01-04 2021-01-05 2021-01-06 2021-01-07 2021-01-08 2021-01-09", cellValues: ["2021-01-03", "2021-01-04", "2021-01-05", "2021-01-06", "2021-01-07", "2021-01-08", "2021-01-09"] },
        { row: 4, startCol: 1, text: "仕事 私用 その他", cellValues: ["仕事", "私用", "その他", "", "", "", ""] }
      ]
    });

    expect(markdown).toContain("2021年1月");
    expect(markdown).toContain("### 日 月 火 水 木 金 土");
    expect(markdown).toContain("2021-01-03 | 2021-01-04 | 2021-01-05 | 2021-01-06 | 2021-01-07 | 2021-01-08 | 2021-01-09");
    expect(markdown).toContain("仕事 | 私用 | その他");
  });
});
