// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootXlsx2mdCore } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixtureDir = path.resolve(__dirname, "./fixtures/table");

const cases = [
  {
    id: "W1",
    file: "table-header-width-unbordered-sample01.xlsx",
    balanced: "A1-D2",
    border: null,
    plannerAware: "A1-D2",
    rows: ["| A | B | C |  |", "| --- | --- | --- | --- |", "| 1 | 2 | 3 | 4 |"]
  },
  {
    id: "W2",
    file: "table-header-width-bordered-sample01.xlsx",
    balanced: "A1-D2",
    border: "A1-D2",
    plannerAware: "A1-D2",
    rows: ["| A | B | C |  |", "| --- | --- | --- | --- |", "| 1 | 2 | 3 | 4 |"]
  },
  {
    id: "W3",
    file: "table-header-width-mixed-borders-sample01.xlsx",
    balanced: "A1-C2",
    border: "A1-C2",
    plannerAware: "A1-C2",
    rows: ["| A | B | C |", "| --- | --- | --- |", "| 1 | 2 | 3 |"]
  },
  {
    id: "M1",
    file: "table-merged-header-unbordered-sample01.xlsx",
    balanced: "A1-C3",
    border: null,
    plannerAware: "A1-C3",
    rows: ["| Group | [←M←] | C |", "| --- | --- | --- |", "| 1 | 2 | 3 |", "| 4 | 5 | 6 |"]
  },
  {
    id: "M2",
    file: "table-merged-header-bordered-sample01.xlsx",
    balanced: "A1-C3",
    border: "A1-C3",
    plannerAware: "A1-C3",
    rows: ["| Group | [←M←] | C |", "| --- | --- | --- |", "| 1 | 2 | 3 |", "| 4 | 5 | 6 |"]
  },
  {
    id: "M3",
    file: "table-merged-header-mixed-borders-sample01.xlsx",
    balanced: "A1-C3",
    border: "A2-C3",
    plannerAware: "A2-C3",
    rows: ["| Group | [←M←] | C |", "| --- | --- | --- |", "| 1 | 2 | 3 |", "| 4 | 5 | 6 |"]
  }
];

function readFixture(fileName) {
  const bytes = readFileSync(path.join(fixtureDir, fileName));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function tableLines(markdown) {
  return markdown.split("\n").filter((line) => line.startsWith("|"));
}

describe("xlsx2md merged borderless table header regression", () => {
  it("parses all six issue fixtures and preserves their exact table output", async () => {
    const api = bootXlsx2mdCore(__dirname);

    for (const testCase of cases) {
      const workbook = await api.parseWorkbook(readFixture(testCase.file), testCase.file);
      const sheet = workbook.sheets[0];
      if (testCase.id === "W2") {
        expect(sheet.cells.find((cell) => cell.address === "D1")).toBeUndefined();
        expect(sheet.cells.map((cell) => cell.address).sort()).toEqual([
          "A1", "A2", "B1", "B2", "C1", "C2", "D2"
        ]);
        for (const cell of sheet.cells) {
          expect(cell.borders).toEqual({ top: true, bottom: true, left: true, right: true });
        }
      }
      expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe(testCase.id.startsWith("W") ? "A" : "Group");
      if (testCase.id.startsWith("M")) {
        expect(sheet.merges.map((merge) => merge.ref)).toContain("A1:B1");
      }

      const output = api.convertWorkbookToMarkdownFiles(workbook, {
        tableDetectionMode: "balanced",
        treatFirstRowAsHeader: true,
        trimText: true,
        removeEmptyRows: true,
        removeEmptyColumns: true
      })[0];
      expect(output.summary.tables).toBe(1);
      expect(output.summary.tableScores.map((detail) => detail.range)).toEqual([testCase.balanced]);
      expect(tableLines(output.markdown)).toEqual(testCase.rows);

      if (testCase.id === "W3") {
        expect(output.markdown.split("\n").filter((line) => line.trim() === "4")).toHaveLength(1);
        expect(tableLines(output.markdown).join("\n")).not.toContain("| 4 |");
      }
      if (testCase.id === "M3") {
        expect(output.markdown.match(/Group/gu)).toHaveLength(1);
      }
    }
  });

  it("keeps the documented mode matrix for all issue fixtures", async () => {
    const api = bootXlsx2mdCore(__dirname);
    const modes = ["balanced", "border", "planner-aware"];

    for (const testCase of cases) {
      const workbook = await api.parseWorkbook(readFixture(testCase.file), testCase.file);
      for (const mode of modes) {
        const output = api.convertWorkbookToMarkdownFiles(workbook, {
          tableDetectionMode: mode,
          treatFirstRowAsHeader: true,
          trimText: true,
          removeEmptyRows: true,
          removeEmptyColumns: true
        })[0];
        const expectedRange = mode === "balanced"
          ? testCase.balanced
          : mode === "border"
            ? testCase.border
            : testCase.plannerAware;
        expect(output.summary.tableScores.map((detail) => detail.range)).toEqual(expectedRange ? [expectedRange] : []);
        if (testCase.id === "W2") {
          expect(tableLines(output.markdown)).toEqual(testCase.rows);
        }
      }
    }
  });

  it("renders the merged header as data when requested and retains empty rows and columns", async () => {
    const api = bootXlsx2mdCore(__dirname);
    const fileName = "table-merged-header-mixed-borders-sample01.xlsx";
    const workbook = await api.parseWorkbook(readFixture(fileName), fileName);

    const noHeader = api.convertWorkbookToMarkdownFiles(workbook, {
      tableDetectionMode: "balanced",
      treatFirstRowAsHeader: false,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    })[0];
    expect(tableLines(noHeader.markdown)).toEqual([
      "|  |  |  |",
      "| --- | --- | --- |",
      "| Group | [←M←] | C |",
      "| 1 | 2 | 3 |",
      "| 4 | 5 | 6 |"
    ]);

    const preserveEmpty = api.convertWorkbookToMarkdownFiles(workbook, {
      tableDetectionMode: "balanced",
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: false,
      removeEmptyColumns: false
    })[0];
    expect(preserveEmpty.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-C3"]);
    expect(tableLines(preserveEmpty.markdown).every((line) => line.split("|").length === 5)).toBe(true);
  });
});
