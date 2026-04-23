// @vitest-environment jsdom

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootRegisteredModule } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootTableDetector() {
  return bootRegisteredModule(__dirname, [
    "src/js/border-grid.js",
    "src/js/table-detector.js"
  ], "tableDetector");
}

function createCell(row, col, outputValue, borders = {}) {
  return {
    row,
    col,
    outputValue,
    borders: {
      top: false,
      bottom: false,
      left: false,
      right: false,
      ...borders
    }
  };
}

function buildCellMap(sheet) {
  const cellMap = new Map();
  for (const cell of sheet.cells) {
    cellMap.set(`${cell.row}:${cell.col}`, cell);
  }
  return cellMap;
}

describe("xlsx2md table detector", () => {
  it("collects seed cells from values or borders", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, ""),
        createCell(1, 2, "項目"),
        createCell(2, 1, "", { bottom: true })
      ],
      merges: []
    };

    const seeds = api.collectTableSeedCells(sheet);

    expect(seeds.map((cell) => `${cell.row}:${cell.col}`)).toEqual(["1:2", "2:1"]);
  });

  it("collects border seed cells separately from value-only cells", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "タイトル"),
        createCell(2, 1, "項目", { bottom: true }),
        createCell(2, 2, "値", { bottom: true })
      ],
      merges: []
    };

    const seeds = api.collectBorderSeedCells(sheet);

    expect(seeds.map((cell) => `${cell.row}:${cell.col}`)).toEqual(["2:1", "2:2"]);
  });

  it("breaks a bordered table before a following borderless note row", () => {
    const api = bootTableDetector();
    const cellMap = new Map([
      ["1:1", createCell(1, 1, "項番", { top: true, bottom: true, left: true })],
      ["1:2", createCell(1, 2, "名称", { top: true, bottom: true, right: true })],
      ["2:1", createCell(2, 1, "1", { bottom: true, left: true })],
      ["2:2", createCell(2, 2, "コード", { bottom: true, right: true })],
      ["3:1", createCell(3, 1, "※注記")],
      ["3:2", createCell(3, 2, "")]
    ]);

    const trimmed = api.trimTableCandidateBounds(cellMap, {
      startRow: 1,
      startCol: 1,
      endRow: 3,
      endCol: 2
    });

    expect(trimmed).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2
    });
  });

  it("normalizes candidate matrices by trimming empty rows and columns while keeping merge markers non-meaningful", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "Header"),
        createCell(1, 2, ""),
        createCell(1, 3, ""),
        createCell(2, 1, "Value"),
        createCell(2, 2, "A"),
        createCell(2, 3, ""),
        createCell(3, 1, ""),
        createCell(3, 2, ""),
        createCell(3, 3, "")
      ],
      merges: [
        { startRow: 1, startCol: 1, endRow: 1, endCol: 2, ref: "A1:B1" }
      ]
    };

    const matrix = api.matrixFromCandidate(
      sheet,
      { startRow: 1, startCol: 1, endRow: 3, endCol: 3, score: 5, reasonSummary: [] },
      { trimText: true, removeEmptyRows: true, removeEmptyColumns: true },
      buildCellMap,
      (cell) => cell?.outputValue || ""
    );

    expect(matrix).toEqual([
      ["Header", "[←M←]"],
      ["Value", "A"]
    ]);
  });

  it("detects a bordered dense grid as a table candidate", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "項番", { top: true, bottom: true, left: true }),
        createCell(1, 2, "名称", { top: true, bottom: true, right: true }),
        createCell(2, 1, "1", { bottom: true, left: true }),
        createCell(2, 2, "コード", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap);

    expect(candidates).toHaveLength(1);
    expect(candidates[0]).toMatchObject({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2
    });
    expect(candidates[0].score).toBeGreaterThanOrEqual(api.defaultTableScoreWeights.threshold);
  });

  it("prunes a wider candidate that redundantly contains a tighter table candidate", () => {
    const api = bootTableDetector();

    const pruned = api.pruneRedundantCandidates([
      { startRow: 2, startCol: 1, endRow: 10, endCol: 12, score: 7, reasonSummary: [] },
      { startRow: 2, startCol: 1, endRow: 10, endCol: 7, score: 8, reasonSummary: [] }
    ]);

    expect(pruned).toEqual([
      { startRow: 2, startCol: 1, endRow: 10, endCol: 7, score: 8, reasonSummary: [] }
    ]);
  });

  it("prefers bordered candidates and excludes a borderless title row above the table", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "商品別計算表"),
        createCell(2, 1, "商CO", { top: true, bottom: true, left: true }),
        createCell(2, 2, "商品名", { top: true, bottom: true }),
        createCell(2, 3, "仕入数", { top: true, bottom: true, right: true }),
        createCell(3, 1, "101", { bottom: true, left: true }),
        createCell(3, 2, "商品A", { bottom: true }),
        createCell(3, 3, "693", { bottom: true, right: true })
      ],
      merges: [
        { startRow: 1, startCol: 1, endRow: 1, endCol: 3, ref: "A1:C1" }
      ]
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap);

    expect(candidates).toHaveLength(1);
    expect(candidates[0]).toMatchObject({
      startRow: 2,
      startCol: 1,
      endRow: 3,
      endCol: 3
    });
  });

  it("does not keep a wide fallback candidate when multiple bordered tables fill most of the area", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "表1"),
        createCell(1, 5, "表2"),
        createCell(2, 1, "項番", { top: true, bottom: true, left: true }),
        createCell(2, 2, "名称", { top: true, bottom: true }),
        createCell(2, 3, "値", { top: true, bottom: true, right: true }),
        createCell(2, 5, "項番", { top: true, bottom: true, left: true }),
        createCell(2, 6, "名称", { top: true, bottom: true }),
        createCell(2, 7, "値", { top: true, bottom: true, right: true }),
        createCell(3, 1, "1", { bottom: true, left: true }),
        createCell(3, 2, "A", { bottom: true }),
        createCell(3, 3, "100", { bottom: true, right: true }),
        createCell(3, 5, "1", { bottom: true, left: true }),
        createCell(3, 6, "B", { bottom: true }),
        createCell(3, 7, "200", { bottom: true, right: true }),
        createCell(4, 1, "表3"),
        createCell(4, 5, "表4"),
        createCell(5, 1, "項番", { top: true, bottom: true, left: true }),
        createCell(5, 2, "名称", { top: true, bottom: true }),
        createCell(5, 3, "値", { top: true, bottom: true, right: true }),
        createCell(5, 5, "項番", { top: true, bottom: true, left: true }),
        createCell(5, 6, "名称", { top: true, bottom: true }),
        createCell(5, 7, "値", { top: true, bottom: true, right: true }),
        createCell(6, 1, "1", { bottom: true, left: true }),
        createCell(6, 2, "C", { bottom: true }),
        createCell(6, 3, "300", { bottom: true, right: true }),
        createCell(6, 5, "1", { bottom: true, left: true }),
        createCell(6, 6, "D", { bottom: true }),
        createCell(6, 7, "400", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap);

    expect(candidates.map((candidate) => ({
      startRow: candidate.startRow,
      startCol: candidate.startCol,
      endRow: candidate.endRow,
      endCol: candidate.endCol
    }))).toEqual([
      { startRow: 2, startCol: 1, endRow: 3, endCol: 3 },
      { startRow: 2, startCol: 5, endRow: 3, endCol: 7 },
      { startRow: 5, startCol: 1, endRow: 6, endCol: 3 },
      { startRow: 5, startCol: 5, endRow: 6, endCol: 7 }
    ]);
  });

  it("does not treat a wide sparse merge-heavy bordered form block as a table", () => {
    const api = bootTableDetector();
    const cells = [];
    for (let row = 1; row <= 6; row += 1) {
      for (let col = 1; col <= 12; col += 1) {
        let value = "";
        if (col === 1) value = `label${row}`;
        if (row === 2 && col === 8) value = "開始日時";
        if (row === 4 && col === 8) value = "終了日時";
        cells.push(createCell(row, col, value, { top: true, bottom: true, left: true, right: true }));
      }
    }
    const sheet = {
      cells,
      merges: [
        { startRow: 1, startCol: 2, endRow: 1, endCol: 12, ref: "B1:L1" },
        { startRow: 2, startCol: 2, endRow: 2, endCol: 7, ref: "B2:G2" },
        { startRow: 2, startCol: 8, endRow: 2, endCol: 12, ref: "H2:L2" },
        { startRow: 3, startCol: 2, endRow: 3, endCol: 12, ref: "B3:L3" },
        { startRow: 4, startCol: 2, endRow: 4, endCol: 7, ref: "B4:G4" },
        { startRow: 4, startCol: 8, endRow: 4, endCol: 12, ref: "H4:L4" },
        { startRow: 5, startCol: 2, endRow: 5, endCol: 12, ref: "B5:L5" },
        { startRow: 6, startCol: 2, endRow: 6, endCol: 12, ref: "B6:L6" }
      ]
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap);

    expect(candidates).toHaveLength(0);
  });

  it("planner-aware drops calendar-like narrow table candidates when many columns line up in the same band", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "月", { top: true, bottom: true, left: true }),
        createCell(1, 2, "予定", { top: true, bottom: true, right: true }),
        createCell(2, 1, "1", { bottom: true, left: true }),
        createCell(2, 2, "A", { bottom: true, right: true }),
        createCell(3, 1, "2", { bottom: true, left: true }),
        createCell(3, 2, "B", { bottom: true, right: true }),
        createCell(4, 1, "3", { bottom: true, left: true }),
        createCell(4, 2, "C", { bottom: true, right: true }),

        createCell(1, 4, "火", { top: true, bottom: true, left: true }),
        createCell(1, 5, "予定", { top: true, bottom: true, right: true }),
        createCell(2, 4, "1", { bottom: true, left: true }),
        createCell(2, 5, "D", { bottom: true, right: true }),
        createCell(3, 4, "2", { bottom: true, left: true }),
        createCell(3, 5, "E", { bottom: true, right: true }),
        createCell(4, 4, "3", { bottom: true, left: true }),
        createCell(4, 5, "F", { bottom: true, right: true }),

        createCell(1, 7, "水", { top: true, bottom: true, left: true }),
        createCell(1, 8, "予定", { top: true, bottom: true, right: true }),
        createCell(2, 7, "1", { bottom: true, left: true }),
        createCell(2, 8, "G", { bottom: true, right: true }),
        createCell(3, 7, "2", { bottom: true, left: true }),
        createCell(3, 8, "H", { bottom: true, right: true }),
        createCell(4, 7, "3", { bottom: true, left: true }),
        createCell(4, 8, "I", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "planner-aware");

    expect(candidates).toHaveLength(0);
  });

  it("does not treat a tiny merged label stub as a bordered table", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(4, 18, "金曜日", { top: true, bottom: true, left: true }),
        createCell(4, 19, "", { top: true, bottom: true, right: true })
      ],
      merges: [
        { startRow: 4, startCol: 18, endRow: 4, endCol: 19, ref: "R4:S4" }
      ]
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "balanced");

    expect(candidates).toHaveLength(0);
  });

  it("planner-aware does not keep a huge fallback candidate for a merge-heavy mixed layout sheet", () => {
    const api = bootTableDetector();
    const cells = [];
    for (let row = 1; row <= 24; row += 1) {
      for (let col = 1; col <= 9; col += 1) {
        let value = "";
        if (row <= 6 && col === 1) value = `見出し${row}`;
        if (row >= 7 && row <= 24) value = `${row}-${col}`;
        cells.push(createCell(row, col, value));
      }
    }
    const sheet = {
      cells,
      merges: [
        { startRow: 1, startCol: 1, endRow: 1, endCol: 9, ref: "A1:I1" },
        { startRow: 2, startCol: 2, endRow: 2, endCol: 4, ref: "B2:D2" },
        { startRow: 2, startCol: 5, endRow: 2, endCol: 7, ref: "E2:G2" },
        { startRow: 3, startCol: 2, endRow: 3, endCol: 4, ref: "B3:D3" }
      ]
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "planner-aware");

    expect(candidates).toHaveLength(0);
  });

  it("currently detects a dense borderless 2x2 block as a table candidate", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "項目"),
        createCell(1, 2, "値"),
        createCell(2, 1, "A"),
        createCell(2, 2, "100")
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap);

    expect(candidates).toHaveLength(1);
    expect(candidates[0]).toMatchObject({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2
    });
    expect(candidates[0].reasonSummary).toContain("High density (+2)");
  });

  it("balanced mode keeps dense borderless blocks that are currently detected via fallback seeds", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "項目"),
        createCell(1, 2, "値"),
        createCell(2, 1, "A"),
        createCell(2, 2, "100")
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "balanced");

    expect(candidates).toHaveLength(1);
  });

  it("border mode excludes dense borderless blocks that are only detected by fallback seeds", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "項目"),
        createCell(1, 2, "値"),
        createCell(2, 1, "A"),
        createCell(2, 2, "100")
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "border");

    expect(candidates).toHaveLength(0);
  });

  it("border mode still keeps bordered 2x2 tables", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "項番", { top: true, bottom: true, left: true }),
        createCell(1, 2, "名称", { top: true, bottom: true, right: true }),
        createCell(2, 1, "1", { bottom: true, left: true }),
        createCell(2, 2, "コード", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "border");

    expect(candidates).toHaveLength(1);
    expect(candidates[0]).toMatchObject({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2
    });
  });

  it("border mode still detects an inner bordered table that is not border-connected to an outer frame block", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "設計書", { top: true, bottom: true, left: true }),
        createCell(1, 2, "", { top: true, bottom: true }),
        createCell(1, 3, "", { top: true, bottom: true, right: true }),
        createCell(2, 1, "見出し", { left: true }),
        createCell(2, 3, "", { right: true }),
        createCell(3, 1, "", { left: true }),
        createCell(3, 3, "", { right: true }),
        createCell(4, 1, "", { left: true, bottom: true }),
        createCell(4, 2, "", { bottom: true }),
        createCell(4, 3, "", { right: true, bottom: true }),

        createCell(6, 1, "項目", { top: true, bottom: true, left: true }),
        createCell(6, 2, "値", { top: true, bottom: true, right: true }),
        createCell(7, 1, "A", { bottom: true, left: true }),
        createCell(7, 2, "100", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "border");

    expect(candidates).toEqual(expect.arrayContaining([
      expect.objectContaining({
        startRow: 6,
        startCol: 1,
        endRow: 7,
        endCol: 2
      })
    ]));
  });

  it("keeps an inner bordered table even when it sits inside the bounds of a separate outer frame block", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "設計書", { top: true, bottom: true, left: true }),
        createCell(1, 2, "", { top: true, bottom: true }),
        createCell(1, 3, "", { top: true, bottom: true }),
        createCell(1, 4, "", { top: true, bottom: true, right: true }),
        createCell(2, 1, "見出し", { left: true }),
        createCell(2, 4, "", { right: true }),
        createCell(3, 1, "", { left: true }),
        createCell(3, 4, "", { right: true }),
        createCell(4, 1, "", { left: true }),
        createCell(4, 4, "", { right: true }),
        createCell(5, 1, "", { left: true }),
        createCell(5, 4, "", { right: true }),
        createCell(6, 1, "", { left: true, bottom: true }),
        createCell(6, 2, "", { bottom: true }),
        createCell(6, 3, "", { bottom: true }),
        createCell(6, 4, "", { right: true, bottom: true }),

        createCell(4, 2, "項目", { top: true, bottom: true, left: true }),
        createCell(4, 3, "値", { top: true, bottom: true, right: true }),
        createCell(5, 2, "A", { bottom: true, left: true }),
        createCell(5, 3, "100", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap, undefined, "border");

    expect(candidates).toEqual(expect.arrayContaining([
      expect.objectContaining({
        startRow: 4,
        startCol: 2,
        endRow: 5,
        endCol: 3
      })
    ]));
  });
});
