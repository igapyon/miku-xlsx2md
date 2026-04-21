// @vitest-environment jsdom

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootSheetMarkdownModule } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootSheetMarkdown() {
  return bootSheetMarkdownModule(__dirname);
}

function createDeps(overrides = {}) {
  return {
    renderNarrativeBlock: (block) => `### ${block.lines[0]}`,
    detectTableCandidates: () => [],
    matrixFromCandidate: () => [],
    renderMarkdownTable: (rows) => rows.map((row) => `| ${row.join(" | ")} |`).join("\n"),
    createOutputFileName: (_workbookName, sheetIndex, sheetName, _outputMode = "display", formattingMode = "plain") => (
      `${sheetIndex}_${sheetName}${formattingMode === "plain" ? "" : `_${formattingMode}`}.md`
    ),
    extractShapeBlocks: () => [],
    renderHierarchicalRawEntries: () => [],
    parseCellAddress: (address) => {
      const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
      if (!match) return { row: 0, col: 0 };
      let col = 0;
      for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
      return { row: Number(match[2]), col };
    },
    formatRange: (startRow, startCol, endRow, endCol) => `${startRow}:${startCol}-${endRow}:${endCol}`,
    colToLetters: (col) => String.fromCharCode(64 + col),
    normalizeMarkdownText: (text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\s+/g, " ").trim(),
    defaultCellWidthEmu: 1,
    defaultCellHeightEmu: 1,
    shapeBlockGapXEmu: 1,
    shapeBlockGapYEmu: 1,
    ...overrides
  };
}

describe("xlsx2md sheet markdown", () => {
  it("builds cell maps and narrative row segments", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const sheet = {
      cells: [
        { row: 1, col: 1, outputValue: "A", rawValue: "A" },
        { row: 1, col: 2, outputValue: "B", rawValue: "B" },
        { row: 1, col: 8, outputValue: "C", rawValue: "C" }
      ]
    };

    expect(api.buildCellMap(sheet).get("1:2")?.outputValue).toBe("B");
    expect(api.splitNarrativeRowSegments(sheet.cells, {})).toEqual([
      { startCol: 1, values: ["A", "B"] },
      { startCol: 8, values: ["C"] }
    ]);
  });

  it("collects narrative cells by row and builds narrative items in row order", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      cells: [
        { row: 2, col: 4, outputValue: "Tail", rawValue: "Tail" },
        { row: 1, col: 1, outputValue: "Head", rawValue: "Head" },
        { row: 2, col: 1, outputValue: "Body", rawValue: "Body" },
        { row: 5, col: 1, outputValue: "TableCell", rawValue: "TableCell" }
      ]
    };
    const tables = [{ startRow: 5, startCol: 1, endRow: 5, endCol: 1 }];

    const rowMap = api.collectNarrativeCellsByRow(sheet, tables);
    const items = api.buildNarrativeItems(workbook, sheet, tables, {});

    expect(Array.from(rowMap.keys())).toEqual([2, 1]);
    expect(items).toEqual([
      { row: 1, startCol: 1, text: "Head", cellValues: ["Head"] },
      { row: 2, startCol: 1, text: "Body Tail", cellValues: ["Body", "Tail"] }
    ]);
  });

  it("extracts narrative blocks outside detected tables", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      cells: [
        { row: 1, col: 1, outputValue: "Heading", rawValue: "Heading" },
        { row: 2, col: 2, outputValue: "Detail", rawValue: "Detail" },
        { row: 5, col: 1, outputValue: "TableCell", rawValue: "TableCell" }
      ],
      charts: [],
      images: []
    };
    const tables = [{ startRow: 5, startCol: 1, endRow: 5, endCol: 1 }];

    expect(api.extractNarrativeBlocks(workbook, sheet, tables, {})).toHaveLength(1);
    expect(api.extractNarrativeBlocks(workbook, sheet, tables, {})[0].lines).toEqual(["Heading", "Detail"]);
  });

  it("groups nearby section anchors and splits distant ones", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const sheet = {
      cells: [],
      images: [
        { anchor: "B2", filename: "img1.png", path: "assets/img1.png", data: new Uint8Array([1]) },
        { anchor: "D5", filename: "img2.png", path: "assets/img2.png", data: new Uint8Array([2]) },
        { anchor: "A12", filename: "img3.png", path: "assets/img3.png", data: new Uint8Array([3]) }
      ],
      charts: [
        { anchor: "C13", title: "Chart", chartType: "bar", series: [] }
      ]
    };
    const tables = [
      { startRow: 4, startCol: 1, endRow: 6, endCol: 3, score: 1, reasonSummary: ["test"] }
    ];
    const narrativeBlocks = [
      {
        startRow: 1,
        startCol: 1,
        endRow: 1,
        lines: ["Intro"],
        items: [{ row: 1, startCol: 1, text: "Intro", cellValues: ["Intro"] }]
      }
    ];

    expect(api.extractSectionBlocks(sheet, tables, narrativeBlocks)).toEqual([
      { startRow: 1, startCol: 1, endRow: 6, endCol: 4 },
      { startRow: 12, startCol: 1, endRow: 13, endCol: 3 }
    ]);
  });

  it("sorts and renders grouped sections with separators only between blocks", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const sections = [
      { sortRow: 12, sortCol: 1, markdown: "Third\n", kind: "narrative" },
      { sortRow: 2, sortCol: 3, markdown: "Second\n", kind: "table" },
      { sortRow: 2, sortCol: 1, markdown: "First\n", kind: "narrative" }
    ];
    const sectionBlocks = [
      { startRow: 1, startCol: 1, endRow: 5, endCol: 5 },
      { startRow: 10, startCol: 1, endRow: 15, endCol: 5 }
    ];

    api.sortContentSections(sections);
    const grouped = api.createGroupedSections(sectionBlocks, sections);

    expect(grouped).toEqual([
      {
        block: { startRow: 1, startCol: 1, endRow: 5, endCol: 5 },
        entries: [
          { sortRow: 2, sortCol: 1, markdown: "First\n", kind: "narrative" },
          { sortRow: 2, sortCol: 3, markdown: "Second\n", kind: "table" }
        ]
      },
      {
        block: { startRow: 10, startCol: 1, endRow: 15, endCol: 5 },
        entries: [
          { sortRow: 12, sortCol: 1, markdown: "Third\n", kind: "narrative" }
        ]
      }
    ]);
    expect(api.renderGroupedSectionBody(grouped)).toBe("First\n\nSecond\n\n---\n\nThird");
  });

  it("renders image and chart sections with stable headings and metadata", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());

    const imageSection = api.renderImageSection({
      images: [
        { anchor: "B2", filename: "image_001.png", path: "assets/Sheet1/image_001.png", data: new Uint8Array([1]) }
      ]
    });
    const chartSection = api.renderChartSection([
      {
        anchor: "C4",
        title: "",
        chartType: "line",
        series: [
          { name: "Sales", categoriesRef: "Sheet1!A2:A5", valuesRef: "Sheet1!B2:B5", axis: "secondary" }
        ]
      }
    ]);

    expect(imageSection).toContain("### Image: 001 (B2)");
    expect(imageSection).toContain("- File: assets/Sheet1/image_001.png");
    expect(imageSection).toContain("![image_001.png](assets/Sheet1/image_001.png)");
    expect(chartSection).toContain("### Chart: 001 (C4)");
    expect(chartSection).toContain("- Title: (none)");
    expect(chartSection).toContain("- Type: line");
    expect(chartSection).toContain("- Series:");
    expect(chartSection).toContain("  - Sales");
    expect(chartSection).toContain("    - Axis: secondary");
    expect(chartSection).toContain("    - categories: Sheet1!A2:A5");
    expect(chartSection).toContain("    - values: Sheet1!B2:B5");
  });

  it("renders ungrouped shapes after grouped shape blocks", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderHierarchicalRawEntries: () => ["- kind: rect"]
    }));
    const shapes = [
      {
        anchor: "B3",
        rawEntries: [{ key: "kind", value: "rect" }],
        svgFilename: null,
        svgPath: null,
        svgData: null
      },
      {
        anchor: "E8",
        rawEntries: [{ key: "kind", value: "rect" }],
        svgFilename: "shape_002.svg",
        svgPath: "assets/Sheet1/shape_002.svg",
        svgData: new Uint8Array([2])
      }
    ];
    const shapeBlocks = [
      { startRow: 3, startCol: 2, endRow: 3, endCol: 2, shapeIndexes: [0] }
    ];

    const section = api.renderShapeSection(shapes, shapeBlocks, true);

    expect(section).toContain("### Shape Block: 001 (3:2-3:2)");
    expect(section).toContain("- Shapes: Shape 001");
    expect(section).toContain("### Ungrouped Shapes");
    expect(section).toContain("#### Shape: 002 (E8)");
    expect(section).toContain("![shape_002.svg](assets/Sheet1/shape_002.svg)");
  });

  it("converts a minimal sheet into markdown", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        { address: "A1", row: 1, col: 1, outputValue: "Hello", rawValue: "Hello", formulaText: "", resolutionStatus: null, resolutionSource: null },
        { address: "B2", row: 2, col: 2, outputValue: "World", rawValue: "World", formulaText: "", resolutionStatus: null, resolutionSource: null }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.fileName).toBe("1_Sheet1.md");
    expect(result.markdown).toContain("# Book: book.xlsx");
    expect(result.markdown).toContain("## Sheet: Sheet1");
    expect(result.summary.narrativeBlocks).toBe(1);
  });

  it("collects render state and uses empty-body fallback text", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const state = api.collectSheetRenderState(workbook, sheet, {});
    const markdown = api.createSheetMarkdownText(workbook, sheet, state);
    const summary = api.createSheetSummary(sheet, state);

    expect(state.body).toBe("");
    expect(state.groupedSections).toEqual([]);
    expect(markdown).toContain("_No extractable body content was found._");
    expect(summary.sections).toBe(0);
    expect(summary.tables).toBe(0);
    expect(summary.narrativeBlocks).toBe(0);
  });

  it("normalizes table detection compatibility aliases in core conversion", () => {
    const detectedModes = [];
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      detectTableCandidates: (_sheet, _buildCellMap, tableDetectionMode) => {
        detectedModes.push(tableDetectionMode);
        return [];
      }
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { tableDetectionMode: "border-priority" });

    expect(detectedModes).toEqual(["border"]);
    expect(result.summary.tableDetectionMode).toBe("border");
  });

  it("normalizes cell line breaks into spaces in plain mode", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        { address: "A1", row: 1, col: 1, outputValue: "Line1\nLine2", rawValue: "Raw1\nRaw2", formulaText: "", resolutionStatus: null, resolutionSource: null }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("Line1 Line2");
    expect(result.markdown).not.toContain("Line1\nLine2");
  });

  it("renders cell line breaks as <br> in github mode", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "Line1\nLine2",
          rawValue: "Raw1\nRaw2",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "github" });

    expect(result.markdown).toContain("Line1<br>Line2");
    expect(result.markdown).not.toContain("Line1 Line2");
  });

  it("omits shape sections when includeShapeDetails is disabled", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      extractShapeBlocks: () => [{ startRow: 3, startCol: 2, endRow: 4, endCol: 3, shapeIndexes: [0] }],
      renderHierarchicalRawEntries: () => ["- kind: rect", "- text: dummy"]
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [],
      merges: [],
      images: [],
      charts: [],
      shapes: [{
        anchor: "B3",
        rawEntries: [{ key: "kind", value: "rect" }],
        svgFilename: null,
        svgPath: null,
        svgData: null
      }]
    };

    const enabled = api.convertSheetToMarkdown(workbook, sheet, {});
    const disabled = api.convertSheetToMarkdown(workbook, sheet, { includeShapeDetails: false });

    expect(enabled.markdown).toContain("### Shape Block: 001 (3:2-4:3)");
    expect(enabled.markdown).toContain("#### Shape: 001 (B3)");
    expect(disabled.markdown).not.toContain("### Shape Block:");
    expect(disabled.markdown).not.toContain("#### Shape:");
  });

  it("keeps a blank line between shape items when svg output is present", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      extractShapeBlocks: () => [],
      renderHierarchicalRawEntries: () => ["- kind: rect"]
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [],
      merges: [],
      images: [],
      charts: [],
      shapes: [{
        anchor: "H3",
        rawEntries: [{ key: "kind", value: "rect" }],
        svgFilename: "shape_001.svg",
        svgPath: "assets/Sheet1/shape_001.svg",
        svgData: new Uint8Array([1])
      }, {
        anchor: "K3",
        rawEntries: [{ key: "kind", value: "rect" }],
        svgFilename: "shape_002.svg",
        svgPath: "assets/Sheet1/shape_002.svg",
        svgData: new Uint8Array([2])
      }]
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("![shape_001.svg](assets/Sheet1/shape_001.svg)\n\n#### Shape: 002 (K3)");
  });

  it("keeps line-start markdown markers literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "# heading",
          rawValue: "# heading",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "- item",
          rawValue: "- item",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\# heading");
    expect(result.markdown).toContain("\\- item");
  });

  it("keeps ordered-list and quote markers literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "1. item",
          rawValue: "1. item",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "> quote",
          rawValue: "> quote",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("1\\. item");
    expect(result.markdown).toContain("&gt; quote");
  });

  it("keeps image-like markdown and code spans literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "![alt](img.png)",
          rawValue: "![alt](img.png)",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "`code`",
          rawValue: "`code`",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\!\\[alt\\]\\(img.png\\)");
    expect(result.markdown).toContain("\\`code\\`");
  });

  it("keeps additional list markers and ampersands literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "+ plus",
          rawValue: "+ plus",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "* star",
          rawValue: "* star",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A3",
          row: 3,
          col: 1,
          outputValue: "a & b",
          rawValue: "a & b",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\+ plus");
    expect(result.markdown).toContain("\\* star");
    expect(result.markdown).toContain("a &amp; b");
  });

  it("shows narrative-vs-table differences for the same markdown-like text", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      detectTableCandidates: () => [{
        startRow: 2,
        startCol: 1,
        endRow: 3,
        endCol: 1,
        score: 1,
        reasonSummary: ["test"]
      }],
      matrixFromCandidate: () => [["`code` ![alt](img.png)"], ["a | b"]],
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "`code` ![alt](img.png)",
          rawValue: "`code` ![alt](img.png)",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "`code` ![alt](img.png)",
          rawValue: "`code` ![alt](img.png)",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A3",
          row: 3,
          col: 1,
          outputValue: "a | b",
          rawValue: "a | b",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\`code\\` \\!\\[alt\\]\\(img.png\\)");
    expect(result.markdown).toContain("| `code` ![alt](img.png) |");
    expect(result.markdown).toContain("| a | b |");
  });

  it("renders external and workbook hyperlinks as markdown links", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = {
      name: "book.xlsx",
      sheets: [
        { name: "Sheet1", index: 1, cells: [], merges: [], images: [], charts: [], shapes: [] },
        { name: "Other Sheet", index: 2, cells: [], merges: [], images: [], charts: [], shapes: [] }
      ]
    };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "Open",
          rawValue: "Open",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null,
          hyperlink: {
            kind: "external",
            target: "https://example.com/",
            location: "",
            tooltip: "",
            display: ""
          }
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "Jump",
          rawValue: "Jump",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null,
          hyperlink: {
            kind: "internal",
            target: "'Other Sheet'!C3",
            location: "'Other Sheet'!C3",
            tooltip: "",
            display: ""
          }
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "github" });

    expect(result.markdown).toContain("[Open](https://example.com/)");
    expect(result.markdown).toContain("[Jump](#other-sheet) (Other Sheet!C3)");
  });

  it("suppresses underline markup for hyperlink cells in github mode", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = {
      name: "book.xlsx",
      sheets: [
        { name: "Sheet1", index: 1, cells: [], merges: [], images: [], charts: [], shapes: [] }
      ]
    };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "Linked",
          rawValue: "Linked",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: true },
          richTextRuns: null,
          hyperlink: {
            kind: "external",
            target: "https://example.com/",
            location: "",
            tooltip: "",
            display: ""
          }
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "github" });

    expect(result.markdown).toContain("[Linked](https://example.com/)");
    expect(result.markdown).not.toContain("<ins>Linked</ins>");
  });

  it("uses raw values for raw output mode while preserving hyperlinks", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = {
      name: "book.xlsx",
      sheets: [
        { name: "Sheet1", index: 1, cells: [], merges: [], images: [], charts: [], shapes: [] }
      ]
    };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "Displayed",
          rawValue: "https://raw.example/",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null,
          hyperlink: {
            kind: "external",
            target: "https://example.com/",
            location: "",
            tooltip: "",
            display: ""
          }
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { outputMode: "raw" });

    expect(result.markdown).toContain("[https://raw.example/](https://example.com/)");
    expect(result.markdown).not.toContain("[Displayed](https://example.com/)");
  });

  it("appends raw values only in both output mode when display and raw differ", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "Displayed",
          rawValue: "RawValue",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "SameValue",
          rawValue: "SameValue",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { outputMode: "both" });

    expect(result.markdown).toContain("Displayed [raw=RawValue]");
    expect(result.markdown).toContain("SameValue");
    expect(result.markdown).not.toContain("SameValue [raw=SameValue]");
  });
});
