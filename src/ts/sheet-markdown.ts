/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;

  type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };

  type ParsedCell = {
    address: string;
    row: number;
    col: number;
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: string | null;
    styleIndex: number;
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: {
      bold: boolean;
      italic: boolean;
      strike: boolean;
      underline: boolean;
    };
    richTextRuns: Array<{
      text: string;
      bold: boolean;
      italic: boolean;
      strike: boolean;
      underline: boolean;
    }> | null;
    formulaType: string;
    spillRef: string;
    hyperlink: {
      kind: "external" | "internal";
      target: string;
      location: string;
      tooltip: string;
      display: string;
    } | null;
  };

  type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
  };

  type NarrativeItem = {
    row: number;
    startCol: number;
    text: string;
    cellValues: string[];
  };

  type NarrativeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    lines: string[];
    items: NarrativeItem[];
  };

  type SectionBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };

  type SectionAnchor = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };

  type ParsedImageAsset = {
    anchor: string;
    filename: string;
    path: string;
    data: Uint8Array;
  };

  type ParsedChartAsset = {
    anchor: string;
    title: string;
    chartType: string;
    series: {
      name: string;
      categoriesRef: string;
      valuesRef: string;
      axis: "primary" | "secondary";
    }[];
  };

  type ParsedShapeAsset = {
    anchor: string;
    rawEntries: {
      key: string;
      value: string;
    }[];
    svgFilename: string | null;
    svgPath: string | null;
    svgData: Uint8Array | null;
  };

  type ParsedSheet = {
    name: string;
    index: number;
    cells: ParsedCell[];
    merges: MergeRange[];
    images: ParsedImageAsset[];
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
  };

  type ParsedWorkbook = {
    name: string;
    sheets: ParsedSheet[];
  };

  type MarkdownOptions = {
    treatFirstRowAsHeader?: boolean;
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
    includeShapeDetails?: boolean;
    outputMode?: "display" | "raw" | "both";
    formattingMode?: "plain" | "github";
    tableDetectionMode?: "balanced" | "border";
  };

  type TableCandidate = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    score: number;
    reasonSummary: string[];
  };

  type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
  };

  type ContentSection = {
    sortRow: number;
    sortCol: number;
    markdown: string;
    kind: "narrative" | "table";
    narrativeBlock?: NarrativeBlock;
  };

  type GroupedSection = {
    block: SectionBlock;
    entries: ContentSection[];
  };

  type SheetRenderState = {
    resolvedOptions: Required<MarkdownOptions>;
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
    shapeBlocks: ShapeBlock[];
    tables: TableCandidate[];
    narrativeBlocks: NarrativeBlock[];
    sectionBlocks: SectionBlock[];
    formulaDiagnostics: MarkdownFile["summary"]["formulaDiagnostics"];
    sections: ContentSection[];
    groupedSections: GroupedSection[];
    body: string;
    imageSection: string;
    chartSection: string;
    shapeSection: string;
  };

  type CellMarkdownValues = {
    displayValue: string;
    rawValue: string;
    displayMarkdown: string;
  };

  type CellMarkdownContext = {
    workbook: ParsedWorkbook | null;
    sheet: ParsedSheet | null;
    options: MarkdownOptions;
  };

  type CellMarkdownRenderer = (
    cell: ParsedCell,
    values: CellMarkdownValues,
    context: CellMarkdownContext
  ) => string;

  type NarrativeRowSegment = {
    startCol: number;
    values: string[];
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
      tableScores: Array<{
        range: string;
        score: number;
        reasons: string[];
      }>;
      formulaDiagnostics: Array<{
        address: string;
        formulaText: string;
        status: FormulaResolutionStatus;
        source: string | null;
        outputValue: string;
      }>;
    };
  };

  type SheetMarkdownDeps = {
    renderNarrativeBlock: (block: NarrativeBlock) => string;
    detectTableCandidates: (
      sheet: ParsedSheet,
      buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
      tableDetectionMode?: "balanced" | "border"
    ) => TableCandidate[];
    matrixFromCandidate: (
      sheet: ParsedSheet,
      candidate: TableCandidate,
      options: MarkdownOptions,
      buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
      formatCellForMarkdown: (cell: ParsedCell | undefined, options: MarkdownOptions) => string
    ) => string[][];
    renderMarkdownTable: (rows: string[][], treatFirstRowAsHeader: boolean) => string;
    createOutputFileName: (
      workbookName: string,
      sheetIndex: number,
      sheetName: string,
      outputMode?: "display" | "raw" | "both",
      formattingMode?: "plain" | "github"
    ) => string;
    extractShapeBlocks: (
      shapes: ParsedShapeAsset[],
      options: {
        defaultCellWidthEmu: number;
        defaultCellHeightEmu: number;
        shapeBlockGapXEmu: number;
        shapeBlockGapYEmu: number;
      }
    ) => ShapeBlock[];
    renderHierarchicalRawEntries: (entries: { key: string; value: string }[]) => string[];
    parseCellAddress: (address: string) => { row: number; col: number };
    formatRange: (startRow: number, startCol: number, endRow: number, endCol: number) => string;
    colToLetters: (col: number) => string;
    normalizeMarkdownText?: (text: string) => string;
    defaultCellWidthEmu: number;
    defaultCellHeightEmu: number;
    shapeBlockGapXEmu: number;
    shapeBlockGapYEmu: number;
  };

  function createSheetMarkdownApi(deps: SheetMarkdownDeps) {
    const markdownOptions = requireXlsx2mdMarkdownOptions();
    const richTextRenderer = requireXlsx2mdRichTextRendererModule<ParsedCell>().createRichTextRendererApi({
      normalizeMarkdownText: deps.normalizeMarkdownText
    });

    function buildCellMap(sheet: ParsedSheet): Map<string, ParsedCell> {
      const map = new Map<string, ParsedCell>();
      for (const cell of sheet.cells) {
        map.set(`${cell.row}:${cell.col}`, cell);
      }
      return map;
    }

    function createHeadingFragment(text: string): string {
      return String(text || "")
        .trim()
        .toLowerCase()
        .replace(/<[^>]+>/g, "")
        .replace(/[^\p{L}\p{N}\s_-]+/gu, "")
        .replace(/\s+/g, "-");
    }

    function parseInternalHyperlinkLocation(location: string, currentSheetName: string): { sheetName: string; refText: string } {
      const normalized = String(location || "").trim().replace(/^#/, "");
      if (!normalized) {
        return { sheetName: currentSheetName, refText: "" };
      }
      const match = normalized.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
      if (match) {
        return {
          sheetName: (match[1] || match[2] || currentSheetName).replace(/''/g, "'"),
          refText: (match[3] || "").trim()
        };
      }
      return {
        sheetName: currentSheetName,
        refText: normalized
      };
    }

    function renderHyperlinkMarkdown(
      cell: ParsedCell,
      text: string,
      workbook: ParsedWorkbook | null,
      sheet: ParsedSheet | null,
      options: MarkdownOptions
    ): string {
      const hyperlink = cell.hyperlink;
      const label = String(text || "").trim();
      if (!hyperlink || !label) return text;
      if (hyperlink.kind === "external") {
        const href = String(hyperlink.target || "").trim();
        return href ? `[${label}](${href})` : label;
      }
      const currentSheetName = sheet?.name || "";
      const { sheetName, refText } = parseInternalHyperlinkLocation(hyperlink.location || hyperlink.target, currentSheetName);
      const traceText = [sheetName, refText].filter(Boolean).join("!");
      const targetSheet = workbook?.sheets.find((entry) => entry.name === sheetName) || null;
      if (!targetSheet || !workbook) {
        return traceText ? `${label} (${traceText})` : label;
      }
      const href = `#${createHeadingFragment(targetSheet.name)}`;
      return traceText && traceText !== targetSheet.name
        ? `[${label}](${href}) (${traceText})`
        : `[${label}](${href})`;
    }

    function createDisplayCellForFormatting(cell: ParsedCell, formattingMode: "plain" | "github"): ParsedCell {
      if (formattingMode !== "github" || !cell.hyperlink) {
        return cell;
      }
      return {
        ...cell,
        textStyle: {
          ...cell.textStyle,
          underline: false
        },
        richTextRuns: cell.richTextRuns?.map((run) => ({
          ...run,
          underline: false
        })) || null
      };
    }

    function createCellMarkdownValues(cell: ParsedCell, formattingMode: "plain" | "github"): CellMarkdownValues {
      const displayCell = createDisplayCellForFormatting(cell, formattingMode);
      return {
        displayValue: richTextRenderer.compactText(String(cell.outputValue || "")),
        rawValue: richTextRenderer.compactText(String(cell.rawValue || "")),
        displayMarkdown: richTextRenderer.renderCellDisplayText(displayCell, formattingMode)
      };
    }

    function renderCellWithHyperlink(cell: ParsedCell, text: string, context: CellMarkdownContext): string {
      return renderHyperlinkMarkdown(cell, text, context.workbook, context.sheet, context.options);
    }

    function renderDisplayModeCell(cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext): string {
      return renderCellWithHyperlink(cell, values.displayMarkdown, context);
    }

    function renderRawModeCell(cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext): string {
      return renderCellWithHyperlink(cell, values.rawValue || values.displayValue, context);
    }

    function renderBothModeCell(cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext): string {
      if (values.rawValue && values.rawValue !== values.displayValue) {
        if (values.displayMarkdown) {
          return `${renderCellWithHyperlink(cell, values.displayMarkdown, context)} [raw=${values.rawValue}]`;
        }
        return `[raw=${values.rawValue}]`;
      }
      return renderCellWithHyperlink(cell, values.displayMarkdown || values.rawValue, context);
    }

    function getCellMarkdownRenderer(outputMode: "display" | "raw" | "both"): CellMarkdownRenderer {
      if (outputMode === "raw") {
        return renderRawModeCell;
      }
      if (outputMode === "both") {
        return renderBothModeCell;
      }
      return renderDisplayModeCell;
    }

    function formatCellForMarkdown(
      cell: ParsedCell | undefined,
      options: MarkdownOptions,
      workbook: ParsedWorkbook | null = null,
      sheet: ParsedSheet | null = null
    ): string {
      if (!cell) return "";
      const resolvedOptions = markdownOptions.resolveMarkdownOptions(options);
      const formattingMode = resolvedOptions.formattingMode;
      const values = createCellMarkdownValues(cell, formattingMode);
      const renderCell = getCellMarkdownRenderer(resolvedOptions.outputMode);
      return renderCell(cell, values, {
        workbook,
        sheet,
        options
      });
    }

    function isCellInAnyTable(row: number, col: number, tables: TableCandidate[]): boolean {
      return tables.some((table) => row >= table.startRow && row <= table.endRow && col >= table.startCol && col <= table.endCol);
    }

    function splitNarrativeRowSegments(
      cells: ParsedCell[],
      options: MarkdownOptions,
      workbook: ParsedWorkbook | null = null,
      sheet: ParsedSheet | null = null
    ): NarrativeRowSegment[] {
      const segments: Array<NarrativeRowSegment & { lastCol: number }> = [];
      let current: { startCol: number; values: string[]; lastCol: number } | null = null;
      for (const cell of cells) {
        const value = formatCellForMarkdown(cell, options, workbook, sheet).trim();
        if (!value) continue;
        if (!current || cell.col - current.lastCol > 4) {
          current = {
            startCol: cell.col,
            values: [value],
            lastCol: cell.col
          };
          segments.push(current);
        } else {
          current.values.push(value);
          current.lastCol = cell.col;
        }
      }
      return segments.map((segment) => ({
        startCol: segment.startCol,
        values: segment.values
      }));
    }

    function collectNarrativeCellsByRow(sheet: ParsedSheet, tables: TableCandidate[]): Map<number, ParsedCell[]> {
      const rowMap = new Map<number, ParsedCell[]>();
      for (const cell of sheet.cells) {
        if (!cell.outputValue) continue;
        if (isCellInAnyTable(cell.row, cell.col, tables)) continue;
        const entries = rowMap.get(cell.row) || [];
        entries.push(cell);
        rowMap.set(cell.row, entries);
      }
      return rowMap;
    }

    function createNarrativeItem(rowNumber: number, segment: NarrativeRowSegment): NarrativeItem {
      const text = segment.values.join(" ").trim();
      return {
        row: rowNumber,
        startCol: segment.startCol,
        text,
        cellValues: segment.values
      };
    }

    function createNarrativeItemsForRow(
      rowNumber: number,
      cells: ParsedCell[],
      options: MarkdownOptions,
      workbook: ParsedWorkbook | null = null,
      sheet: ParsedSheet | null = null
    ): NarrativeItem[] {
      return splitNarrativeRowSegments(cells, options, workbook, sheet)
        .map((segment) => createNarrativeItem(rowNumber, segment))
        .filter((item) => !!item.text);
    }

    function buildNarrativeItems(
      workbook: ParsedWorkbook,
      sheet: ParsedSheet,
      tables: TableCandidate[],
      options: MarkdownOptions = {}
    ): NarrativeItem[] {
      const rowMap = collectNarrativeCellsByRow(sheet, tables);
      const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
      const items: NarrativeItem[] = [];
      for (const rowNumber of rowNumbers) {
        const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
        items.push(...createNarrativeItemsForRow(rowNumber, cells, options, workbook, sheet));
      }
      return items;
    }

    function shouldStartNarrativeBlock(
      current: NarrativeBlock | null,
      rowNumber: number,
      previousRow: number,
      startCol: number
    ): boolean {
      return !current || rowNumber - previousRow > 1 || Math.abs(startCol - current.startCol) > 3;
    }

    function createNarrativeBlockFromItem(item: NarrativeItem): NarrativeBlock {
      return {
        startRow: item.row,
        startCol: item.startCol,
        endRow: item.row,
        lines: [item.text],
        items: [item]
      };
    }

    function appendNarrativeItem(block: NarrativeBlock, item: NarrativeItem): void {
      block.lines.push(item.text);
      block.endRow = item.row;
      block.items.push(item);
    }

    function extractNarrativeBlocks(
      workbook: ParsedWorkbook,
      sheet: ParsedSheet,
      tables: TableCandidate[],
      options: MarkdownOptions = {}
    ): NarrativeBlock[] {
      const items = buildNarrativeItems(workbook, sheet, tables, options);
      const blocks: NarrativeBlock[] = [];
      let current: NarrativeBlock | null = null;
      let previousRow = -100;

      for (const item of items) {
        if (shouldStartNarrativeBlock(current, item.row, previousRow, item.startCol)) {
          current = createNarrativeBlockFromItem(item);
          blocks.push(current);
        } else {
          appendNarrativeItem(current, item);
        }
        previousRow = item.row;
      }

      return blocks;
    }

    function createNarrativeSectionAnchors(narrativeBlocks: NarrativeBlock[]): SectionAnchor[] {
      return narrativeBlocks.map((block) => ({
        startRow: block.startRow,
        startCol: block.startCol,
        endRow: block.endRow,
        endCol: Math.max(block.startCol, ...block.items.map((item) => item.startCol))
      }));
    }

    function createTableSectionAnchors(tables: TableCandidate[]): SectionAnchor[] {
      return tables.map((table) => ({
        startRow: table.startRow,
        startCol: table.startCol,
        endRow: table.endRow,
        endCol: table.endCol
      }));
    }

    function createPointSectionAnchor(address: string): SectionAnchor | null {
      const anchor = deps.parseCellAddress(address);
      if (anchor.row > 0 && anchor.col > 0) {
        return { startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col };
      }
      return null;
    }

    function createImageSectionAnchors(sheet: ParsedSheet): SectionAnchor[] {
      return sheet.images
        .map((image) => createPointSectionAnchor(image.anchor))
        .filter((anchor): anchor is SectionAnchor => !!anchor);
    }

    function createChartSectionAnchors(charts: ParsedChartAsset[]): SectionAnchor[] {
      return charts
        .map((chart) => createPointSectionAnchor(chart.anchor))
        .filter((anchor): anchor is SectionAnchor => !!anchor);
    }

    function createSectionAnchors(sheet: ParsedSheet, tables: TableCandidate[], narrativeBlocks: NarrativeBlock[]): SectionAnchor[] {
      const charts = sheet.charts || [];
      return [
        ...createNarrativeSectionAnchors(narrativeBlocks),
        ...createTableSectionAnchors(tables),
        ...createImageSectionAnchors(sheet),
        ...createChartSectionAnchors(charts)
      ];
    }

    function sortSectionAnchors(anchors: SectionAnchor[]): SectionAnchor[] {
      return anchors.sort((left, right) => {
        if (left.startRow !== right.startRow) return left.startRow - right.startRow;
        return left.startCol - right.startCol;
      });
    }

    function createSectionBlockFromAnchor(anchor: SectionAnchor): SectionBlock {
      return {
        startRow: anchor.startRow,
        startCol: anchor.startCol,
        endRow: anchor.endRow,
        endCol: anchor.endCol
      };
    }

    function shouldStartNewSectionBlock(current: SectionBlock | null, anchor: SectionAnchor, previousEndRow: number): boolean {
      const verticalGapThreshold = 4;
      const gap = anchor.startRow - previousEndRow;
      return !current || gap > verticalGapThreshold;
    }

    function extendSectionBlock(section: SectionBlock, anchor: SectionAnchor): void {
      section.startRow = Math.min(section.startRow, anchor.startRow);
      section.startCol = Math.min(section.startCol, anchor.startCol);
      section.endRow = Math.max(section.endRow, anchor.endRow);
      section.endCol = Math.max(section.endCol, anchor.endCol);
    }

    function extractSectionBlocks(sheet: ParsedSheet, tables: TableCandidate[], narrativeBlocks: NarrativeBlock[]): SectionBlock[] {
      const anchors = sortSectionAnchors(createSectionAnchors(sheet, tables, narrativeBlocks));
      if (anchors.length === 0) {
        return [];
      }

      const sections: SectionBlock[] = [];
      let current: SectionBlock | null = null;
      let previousEndRow = -100;

      for (const anchor of anchors) {
        if (shouldStartNewSectionBlock(current, anchor, previousEndRow)) {
          current = createSectionBlockFromAnchor(anchor);
          sections.push(current);
        } else {
          extendSectionBlock(current, anchor);
        }
        previousEndRow = Math.max(previousEndRow, anchor.endRow);
      }

      return sections;
    }

    function createDefaultShapeSvgFilename(shapeIndex: number): string {
      return `shape_${String(shapeIndex + 1).padStart(3, "0")}.svg`;
    }

    function createImageSectionEntry(image: ParsedImageAsset, index: number): string {
      return [
        `### Image: ${String(index + 1).padStart(3, "0")} (${image.anchor})`,
        `- File: ${image.path}`,
        "",
        `![${image.filename}](${image.path})`
      ].join("\n");
    }

    function createChartSeriesLines(chart: ParsedChartAsset): string[] {
      if (chart.series.length === 0) {
        return [];
      }
      const lines = ["- Series:"];
      for (const series of chart.series) {
        lines.push(`  - ${series.name}`);
        if (series.axis === "secondary") lines.push("    - Axis: secondary");
        if (series.categoriesRef) lines.push(`    - categories: ${series.categoriesRef}`);
        if (series.valuesRef) lines.push(`    - values: ${series.valuesRef}`);
      }
      return lines;
    }

    function createChartSectionEntry(chart: ParsedChartAsset, index: number): string {
      return [
        `### Chart: ${String(index + 1).padStart(3, "0")} (${chart.anchor})`,
        `- Title: ${chart.title || "(none)"}`,
        `- Type: ${chart.chartType}`,
        ...createChartSeriesLines(chart)
      ].join("\n");
    }

    function renderShapeDetails(shape: ParsedShapeAsset, shapeIndex: number): string {
      const lines = [
        `#### Shape: ${String(shapeIndex + 1).padStart(3, "0")} (${shape.anchor})`,
        ...deps.renderHierarchicalRawEntries(shape.rawEntries)
      ];
      if (shape.svgPath) {
        lines.push(`- SVG: ${shape.svgPath}`);
        lines.push("");
        lines.push(`![${shape.svgFilename || createDefaultShapeSvgFilename(shapeIndex)}](${shape.svgPath})`);
      }
      return lines.join("\n");
    }

    function createShapeBlockSummaryLine(shapeIndexes: number[]): string {
      return shapeIndexes.map((shapeIndex) => `Shape ${String(shapeIndex + 1).padStart(3, "0")}`).join(", ");
    }

    function createShapeBlockEntry(
      block: ShapeBlock,
      blockIndex: number,
      shapes: ParsedShapeAsset[]
    ): string {
      const shapeDetails = block.shapeIndexes
        .map((shapeIndex) => {
          const shape = shapes[shapeIndex];
          if (!shape) return "";
          return renderShapeDetails(shape, shapeIndex);
        })
        .filter(Boolean)
        .join("\n\n");
      return [
        `### Shape Block: ${String(blockIndex + 1).padStart(3, "0")} (${deps.formatRange(block.startRow, block.startCol, block.endRow, block.endCol)})`,
        `- Shapes: ${createShapeBlockSummaryLine(block.shapeIndexes)}`,
        `- anchorRange: ${deps.colToLetters(block.startCol)}${block.startRow}-${deps.colToLetters(block.endCol)}${block.endRow}`,
        ...(shapeDetails ? ["", shapeDetails] : [])
      ].join("\n");
    }

    function collectUngroupedShapes(shapes: ParsedShapeAsset[], shapeBlocks: ShapeBlock[]): Array<{ shape: ParsedShapeAsset; index: number }> {
      const grouped = new Set(shapeBlocks.flatMap((block) => block.shapeIndexes));
      return shapes
        .map((shape, index) => ({ shape, index }))
        .filter(({ index }) => !grouped.has(index));
    }

    function renderImageSection(sheet: ParsedSheet): string {
      return sheet.images.length > 0
        ? [
          "",
          ...sheet.images.map((image, index) => createImageSectionEntry(image, index))
        ].join("\n\n")
        : "";
    }

    function renderChartSection(charts: ParsedChartAsset[]): string {
      return charts.length > 0
        ? [
          "",
          ...charts.map((chart, index) => createChartSectionEntry(chart, index))
        ].join("\n\n")
        : "";
    }

    function renderShapeSection(shapes: ParsedShapeAsset[], shapeBlocks: ShapeBlock[], includeShapeDetails: boolean): string {
      const ungrouped = collectUngroupedShapes(shapes, shapeBlocks);
      return includeShapeDetails && shapes.length > 0
        ? [
          "",
          ...shapeBlocks.map((block, blockIndex) => createShapeBlockEntry(block, blockIndex, shapes)),
          ...(ungrouped.length === 0
            ? []
            : [
              "",
              "### Ungrouped Shapes",
              "",
              ...ungrouped.map(({ shape, index }) => renderShapeDetails(shape, index))
            ])
        ].join("\n\n")
        : "";
    }

    function createFormulaDiagnostics(sheet: ParsedSheet): MarkdownFile["summary"]["formulaDiagnostics"] {
      return sheet.cells
        .filter((cell) => !!cell.formulaText && cell.resolutionStatus !== null)
        .map((cell) => ({
          address: cell.address,
          formulaText: cell.formulaText,
          status: cell.resolutionStatus,
          source: cell.resolutionSource,
          outputValue: cell.outputValue
        }));
    }

    function createNarrativeSections(narrativeBlocks: NarrativeBlock[]): ContentSection[] {
      return narrativeBlocks.map((block) => ({
        sortRow: block.startRow,
        sortCol: block.startCol,
        markdown: `${deps.renderNarrativeBlock(block)}\n`,
        kind: "narrative",
        narrativeBlock: block
      }));
    }

    function createTableSections(
      workbook: ParsedWorkbook,
      sheet: ParsedSheet,
      tables: TableCandidate[],
      options: MarkdownOptions,
      treatFirstRowAsHeader: boolean
    ): ContentSection[] {
      const sections: ContentSection[] = [];
      let tableCounter = 1;
      for (const table of tables) {
        const rows = deps.matrixFromCandidate(
          sheet,
          table,
          options,
          buildCellMap,
          (cell, tableOptions) => formatCellForMarkdown(cell, tableOptions, workbook, sheet)
        );
        if (rows.length === 0 || rows[0]?.length === 0) continue;
        const tableMarkdown = deps.renderMarkdownTable(rows, treatFirstRowAsHeader);
        sections.push({
          sortRow: table.startRow,
          sortCol: table.startCol,
          markdown: `### Table: ${String(tableCounter).padStart(3, "0")} (${deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol)})\n\n${tableMarkdown}\n`,
          kind: "table"
        });
        tableCounter += 1;
      }
      return sections;
    }

    function sortContentSections(sections: ContentSection[]): ContentSection[] {
      return sections.sort((left, right) => {
        if (left.sortRow !== right.sortRow) return left.sortRow - right.sortRow;
        return left.sortCol - right.sortCol;
      });
    }

    function createFallbackSectionBlock(): SectionBlock {
      return {
        startRow: -1,
        startCol: -1,
        endRow: Number.MAX_SAFE_INTEGER,
        endCol: Number.MAX_SAFE_INTEGER
      };
    }

    function isSectionInsideBlock(section: ContentSection, block: SectionBlock): boolean {
      return section.sortRow >= block.startRow
        && section.sortRow <= block.endRow
        && section.sortCol >= block.startCol
        && section.sortCol <= block.endCol;
    }

    function createGroupedSections(sectionBlocks: SectionBlock[], sections: ContentSection[]): GroupedSection[] {
      const blocks: SectionBlock[] = sectionBlocks.length > 0
        ? sectionBlocks
        : [createFallbackSectionBlock()];
      return blocks.map((block) => ({
        block,
        entries: sections.filter((section) => isSectionInsideBlock(section, block))
      })).filter((group) => group.entries.length > 0);
    }

    function renderGroupedSectionEntries(entries: ContentSection[]): string {
      return entries.map((section) => section.markdown.trimEnd()).join("\n\n").trim();
    }

    function renderGroupedSectionBody(groupedSections: GroupedSection[]): string {
      return groupedSections
        .map((group) => renderGroupedSectionEntries(group.entries))
        .filter(Boolean)
        .join("\n\n---\n\n")
        .trim();
    }

    function collectSheetRenderState(workbook: ParsedWorkbook, sheet: ParsedSheet, options: MarkdownOptions = {}): SheetRenderState {
      const resolvedOptions = markdownOptions.resolveMarkdownOptions(options);
      const charts = sheet.charts || [];
      const shapes = sheet.shapes || [];
      const shapeBlocks = deps.extractShapeBlocks(shapes, {
        defaultCellWidthEmu: deps.defaultCellWidthEmu,
        defaultCellHeightEmu: deps.defaultCellHeightEmu,
        shapeBlockGapXEmu: deps.shapeBlockGapXEmu,
        shapeBlockGapYEmu: deps.shapeBlockGapYEmu
      });
      const treatFirstRowAsHeader = resolvedOptions.treatFirstRowAsHeader;
      const tableDetectionMode = resolvedOptions.tableDetectionMode;
      const tables = deps.detectTableCandidates(sheet, buildCellMap, tableDetectionMode);
      const narrativeBlocks = extractNarrativeBlocks(workbook, sheet, tables, resolvedOptions);
      const sectionBlocks = extractSectionBlocks(sheet, tables, narrativeBlocks);
      const formulaDiagnostics = createFormulaDiagnostics(sheet);
      const sections = [
        ...createNarrativeSections(narrativeBlocks),
        ...createTableSections(workbook, sheet, tables, resolvedOptions, treatFirstRowAsHeader)
      ];
      sortContentSections(sections);
      const groupedSections = createGroupedSections(sectionBlocks, sections);
      const body = renderGroupedSectionBody(groupedSections);
      const imageSection = renderImageSection(sheet);
      const chartSection = renderChartSection(charts);
      const includeShapeDetails = resolvedOptions.includeShapeDetails;
      const shapeSection = renderShapeSection(shapes, shapeBlocks, includeShapeDetails);
      return {
        resolvedOptions,
        charts,
        shapes,
        shapeBlocks,
        tables,
        narrativeBlocks,
        sectionBlocks,
        formulaDiagnostics,
        sections,
        groupedSections,
        body,
        imageSection,
        chartSection,
        shapeSection
      };
    }

    function createSheetMarkdownText(workbook: ParsedWorkbook, sheet: ParsedSheet, state: SheetRenderState): string {
      const markdown = [
        `# Book: ${workbook.name}`,
        "",
        `## Sheet: ${sheet.name}`,
        "",
        state.body || "_No extractable body content was found._",
        state.chartSection,
        state.shapeSection,
        state.imageSection
      ].join("\n");
      return markdown;
    }

    function createSheetSummary(sheet: ParsedSheet, state: SheetRenderState): MarkdownFile["summary"] {
      return {
        outputMode: state.resolvedOptions.outputMode,
        formattingMode: state.resolvedOptions.formattingMode,
        tableDetectionMode: state.resolvedOptions.tableDetectionMode,
        sections: state.sectionBlocks.length,
        tables: state.tables.length,
        narrativeBlocks: state.narrativeBlocks.length,
        merges: sheet.merges.length,
        images: sheet.images.length,
        charts: state.charts.length,
        cells: sheet.cells.length,
        tableScores: state.tables.map((table) => ({
          range: deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol),
          score: table.score,
          reasons: [...table.reasonSummary]
        })),
        formulaDiagnostics: state.formulaDiagnostics
      };
    }

    function convertSheetToMarkdown(workbook: ParsedWorkbook, sheet: ParsedSheet, options: MarkdownOptions = {}): MarkdownFile {
      const state = collectSheetRenderState(workbook, sheet, options);
      const fileName = deps.createOutputFileName(
        workbook.name,
        sheet.index,
        sheet.name,
        state.resolvedOptions.outputMode,
        state.resolvedOptions.formattingMode
      );
      return {
        fileName,
        sheetName: sheet.name,
        markdown: createSheetMarkdownText(workbook, sheet, state),
        summary: createSheetSummary(sheet, state)
      };
    }

    function convertWorkbookToMarkdownFiles(workbook: ParsedWorkbook, options: MarkdownOptions = {}): MarkdownFile[] {
      return workbook.sheets.map((sheet) => convertSheetToMarkdown(workbook, sheet, options));
    }

    return {
      buildCellMap,
      formatCellForMarkdown,
      isCellInAnyTable,
      splitNarrativeRowSegments,
      collectNarrativeCellsByRow,
      buildNarrativeItems,
      extractNarrativeBlocks,
      extractSectionBlocks,
      sortContentSections,
      createGroupedSections,
      renderGroupedSectionBody,
      collectSheetRenderState,
      createSheetMarkdownText,
      createSheetSummary,
      renderImageSection,
      renderChartSection,
      renderShapeSection,
      convertSheetToMarkdown,
      convertWorkbookToMarkdownFiles
    };
  }

  const sheetMarkdownApi = {
    createSheetMarkdownApi
  };

  moduleRegistry.registerModule("sheetMarkdown", sheetMarkdownApi);
})();
