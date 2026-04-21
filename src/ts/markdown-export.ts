/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;

  type FormulaDiagnostic = {
    address: string;
    formulaText: string;
    status: FormulaResolutionStatus;
    source: string | null;
    outputValue: string;
  };

  type TableScoreDetail = {
    range: string;
    score: number;
    reasons: string[];
  };

  type MarkdownFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
      outputMode: "display" | "raw" | "both";
      formattingMode: "plain" | "github";
      tableDetectionMode: "balanced" | "border";
      sections: number;
      tables: number;
      narrativeBlocks: number;
      merges: number;
      images: number;
      charts: number;
      cells: number;
      tableScores: TableScoreDetail[];
      formulaDiagnostics: FormulaDiagnostic[];
    };
  };

  type ExportEntry = { name: string; data: Uint8Array };
  type MarkdownEncoding = "utf-8" | "shift_jis" | "utf-16le" | "utf-16be" | "utf-32le" | "utf-32be";
  type MarkdownBomMode = "off" | "on";
  type MarkdownEncodingOptions = {
    encoding?: MarkdownEncoding | string | null;
    bom?: MarkdownBomMode | string | null;
  };

  type WorkbookLike = {
    name: string;
    sheets: Array<{
      images: Array<{
        path: string;
        data: Uint8Array;
      }>;
      shapes?: Array<{
        svgPath: string | null;
        svgData: Uint8Array | null;
      }>;
    }>;
  };

  const textEncoder = new TextEncoder();
  const zipIoHelper = requireXlsx2mdZipIo();
  const textEncodingHelper = requireXlsx2mdTextEncoding();
  const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();
  const markdownTableEscapeHelper = requireXlsx2mdMarkdownTableEscape();
  const FORMULA_STATUSES = ["resolved", "fallback_formula", "unsupported_external"] as const;

  function normalizeMarkdownLineBreaks(text: string): string {
    return markdownNormalizeHelper.normalizeMarkdownText(text);
  }

  function escapeMarkdownCell(text: string): string {
    return markdownTableEscapeHelper.escapeMarkdownTableCell(text);
  }

  function createBlankRow(columnCount: number): string[] {
    return new Array(columnCount).fill("");
  }

  function renderMarkdownTable(rows: string[][], treatFirstRowAsHeader: boolean): string {
    if (rows.length === 0) {
      return "";
    }
    const workingRows = rows.map((row) => row.map((cell) => escapeMarkdownCell(cell)));
    if (workingRows.length === 1 && treatFirstRowAsHeader) {
      workingRows.push(createBlankRow(workingRows[0].length));
    }
    const header = treatFirstRowAsHeader ? workingRows[0] : createBlankRow(workingRows[0].length);
    const body = treatFirstRowAsHeader ? workingRows.slice(1) : workingRows;
    const lines = [
      `| ${header.join(" | ")} |`,
      `| ${header.map(() => "---").join(" | ")} |`
    ];
    for (const row of body) {
      lines.push(`| ${row.join(" | ")} |`);
    }
    return lines.join("\n");
  }

  function sanitizeFileNameSegment(value: string, fallback: string): string {
    const normalized = String(value || "").normalize("NFKC");
    const sanitized = normalized
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "_")
      .replace(/[^\p{L}\p{N}._-]+/gu, "_")
      .replace(/_+/g, "_")
      .replace(/^[_ .-]+|[_ .-]+$/g, "");
    return sanitized || fallback;
  }

  function stripWorkbookExtension(workbookName: string): string {
    return String(workbookName || "").replace(/\.xlsx$/i, "");
  }

  function createCombinedMarkdownFileName(workbookName: string): string {
    const baseName = stripWorkbookExtension(String(workbookName || "workbook")) || "workbook";
    return `${baseName}.md`;
  }

  function createExportEntryName(relativePath: string): string {
    return `output/${relativePath}`;
  }

  function createOutputFileName(
    workbookName: string,
    sheetIndex: number,
    sheetName: string,
    outputMode: "display" | "raw" | "both" = "display",
    formattingMode: "plain" | "github" = "plain"
  ): string {
    const bookBase = sanitizeFileNameSegment(stripWorkbookExtension(workbookName), "workbook");
    const safeSheetName = sanitizeFileNameSegment(sheetName, `Sheet${sheetIndex}`);
    void outputMode;
    void formattingMode;
    return `${bookBase}_${String(sheetIndex).padStart(3, "0")}_${safeSheetName}.md`;
  }

  function countFormulaStatuses(formulaDiagnostics: FormulaDiagnostic[]): Record<typeof FORMULA_STATUSES[number], number> {
    const counts = {
      resolved: 0,
      fallback_formula: 0,
      unsupported_external: 0
    };
    for (const item of formulaDiagnostics) {
      if (item.status && item.status in counts) {
        counts[item.status] += 1;
      }
    }
    return counts;
  }

  function createSummaryText(markdownFile: MarkdownFile): string {
    const formulaCounts = countFormulaStatuses(markdownFile.summary.formulaDiagnostics);
    return [
      `Output file: ${markdownFile.fileName}`,
      `Output mode: ${markdownFile.summary.outputMode}`,
      `Formatting mode: ${markdownFile.summary.formattingMode}`,
      `Table detection mode: ${markdownFile.summary.tableDetectionMode}`,
      `Sections: ${markdownFile.summary.sections}`,
      `Tables: ${markdownFile.summary.tables}`,
      `Narrative blocks: ${markdownFile.summary.narrativeBlocks}`,
      `Merged ranges: ${markdownFile.summary.merges}`,
      `Images: ${markdownFile.summary.images}`,
      `Charts: ${markdownFile.summary.charts}`,
      `Analyzed cells: ${markdownFile.summary.cells}`,
      ...FORMULA_STATUSES.map((status) => `Formula ${status}: ${formulaCounts[status]}`),
      ...markdownFile.summary.tableScores.map((detail) => `Table candidate ${detail.range}: score ${detail.score} / ${detail.reasons.join(", ")}`)
    ].join("\n");
  }

  function stripBookHeading(markdown: string, bookHeading: string): string {
    const lines = String(markdown || "").split("\n");
    if (lines[0] === bookHeading) {
      lines.shift();
      while (lines[0] === "") {
        lines.shift();
      }
    }
    return lines.join("\n");
  }

  function createCombinedMarkdownExportFile(workbook: WorkbookLike, markdownFiles: MarkdownFile[]): { fileName: string; content: string } {
    const fileName = createCombinedMarkdownFileName(workbook.name);
    const bookHeading = `# Book: ${String(workbook.name || "workbook.xlsx")}`;
    const content = [
      bookHeading,
      ...markdownFiles
        .map((markdownFile) => stripBookHeading(markdownFile.markdown, bookHeading))
        .filter((markdown) => markdown.trim().length > 0)
    ].join("\n\n");
    return { fileName, content };
  }

  function encodeMarkdownText(text: string, options: MarkdownEncodingOptions = {}): Uint8Array {
    return textEncodingHelper.encodeText(text, options);
  }

  function createCombinedMarkdownExportPayload(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownEncodingOptions = {}
  ): { fileName: string; content: string; data: Uint8Array; mimeType: string } {
    const combined = createCombinedMarkdownExportFile(workbook, markdownFiles);
    return {
      ...combined,
      data: encodeMarkdownText(`${combined.content}\n`, options),
      mimeType: textEncodingHelper.createTextMimeType(options)
    };
  }

  function createMarkdownExportEntry(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownEncodingOptions = {}
  ): ExportEntry | null {
    if (markdownFiles.length === 0) {
      return null;
    }
    const combined = createCombinedMarkdownExportPayload(workbook, markdownFiles, options);
    return {
      name: createExportEntryName(combined.fileName),
      data: combined.data
    };
  }

  function createAssetExportEntries(workbook: WorkbookLike): ExportEntry[] {
    const entries: ExportEntry[] = [];
    for (const sheet of workbook.sheets) {
      for (const image of sheet.images) {
        entries.push({
          name: createExportEntryName(image.path),
          data: image.data
        });
      }
      for (const shape of sheet.shapes || []) {
        if (!shape.svgPath || !shape.svgData) continue;
        entries.push({
          name: createExportEntryName(shape.svgPath),
          data: shape.svgData
        });
      }
    }
    return entries;
  }

  function createExportEntries(workbook: WorkbookLike, markdownFiles: MarkdownFile[], options: MarkdownEncodingOptions = {}): ExportEntry[] {
    const entries = createAssetExportEntries(workbook);
    const markdownEntry = createMarkdownExportEntry(workbook, markdownFiles, options);
    if (markdownEntry) {
      entries.unshift(markdownEntry);
    }
    return entries;
  }

  function createWorkbookExportArchive(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownEncodingOptions = {}
  ): Uint8Array {
    return zipIoHelper.createStoredZip(createExportEntries(workbook, markdownFiles, options));
  }

  const markdownExportApi = {
    encodeMarkdownText,
    createCombinedMarkdownExportPayload,
    escapeMarkdownCell,
    renderMarkdownTable,
    sanitizeFileNameSegment,
    stripWorkbookExtension,
    createCombinedMarkdownFileName,
    createExportEntryName,
    createOutputFileName,
    createSummaryText,
    createCombinedMarkdownExportFile,
    createMarkdownExportEntry,
    createAssetExportEntries,
    createExportEntries,
    createWorkbookExportArchive,
    normalizeMarkdownLineBreaks,
    textEncoder
  };

  moduleRegistry.registerModule("markdownExport", markdownExportApi);
})();
