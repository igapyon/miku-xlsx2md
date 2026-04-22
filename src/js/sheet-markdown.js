/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createSheetMarkdownApi(deps) {
        const markdownOptions = requireXlsx2mdMarkdownOptions();
        const richTextRenderer = requireXlsx2mdRichTextRendererModule().createRichTextRendererApi({
            normalizeMarkdownText: deps.normalizeMarkdownText
        });
        function buildCellMap(sheet) {
            const map = new Map();
            for (const cell of sheet.cells) {
                map.set(`${cell.row}:${cell.col}`, cell);
            }
            return map;
        }
        function createHeadingFragment(text) {
            return String(text || "")
                .trim()
                .toLowerCase()
                .replace(/<[^>]+>/g, "")
                .replace(/[^\p{L}\p{N}\s_-]+/gu, "")
                .replace(/\s+/g, "-");
        }
        function parseInternalHyperlinkLocation(location, currentSheetName) {
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
        function renderHyperlinkMarkdown(cell, text, workbook, sheet, options) {
            const hyperlink = cell.hyperlink;
            const label = String(text || "").trim();
            if (!hyperlink || !label)
                return text;
            if (hyperlink.kind === "external") {
                const href = String(hyperlink.target || "").trim();
                return href ? `[${label}](${href})` : label;
            }
            const currentSheetName = (sheet === null || sheet === void 0 ? void 0 : sheet.name) || "";
            const { sheetName, refText } = parseInternalHyperlinkLocation(hyperlink.location || hyperlink.target, currentSheetName);
            const traceText = [sheetName, refText].filter(Boolean).join("!");
            const targetSheet = (workbook === null || workbook === void 0 ? void 0 : workbook.sheets.find((entry) => entry.name === sheetName)) || null;
            if (!targetSheet || !workbook) {
                return traceText ? `${label} (${traceText})` : label;
            }
            const href = `#${createHeadingFragment(targetSheet.name)}`;
            return traceText && traceText !== targetSheet.name
                ? `[${label}](${href}) (${traceText})`
                : `[${label}](${href})`;
        }
        function createDisplayCellForFormatting(cell, formattingMode) {
            var _a;
            if (formattingMode !== "github" || !cell.hyperlink) {
                return cell;
            }
            return {
                ...cell,
                textStyle: {
                    ...cell.textStyle,
                    underline: false
                },
                richTextRuns: ((_a = cell.richTextRuns) === null || _a === void 0 ? void 0 : _a.map((run) => ({
                    ...run,
                    underline: false
                }))) || null
            };
        }
        function createCellMarkdownValues(cell, formattingMode) {
            const displayCell = createDisplayCellForFormatting(cell, formattingMode);
            return {
                displayValue: richTextRenderer.compactText(String(cell.outputValue || "")),
                rawValue: richTextRenderer.compactText(String(cell.rawValue || "")),
                displayMarkdown: richTextRenderer.renderCellDisplayText(displayCell, formattingMode)
            };
        }
        function renderCellWithHyperlink(cell, text, context) {
            return renderHyperlinkMarkdown(cell, text, context.workbook, context.sheet, context.options);
        }
        function renderDisplayModeCell(cell, values, context) {
            return renderCellWithHyperlink(cell, values.displayMarkdown, context);
        }
        function renderRawModeCell(cell, values, context) {
            return renderCellWithHyperlink(cell, values.rawValue || values.displayValue, context);
        }
        function renderBothModeCell(cell, values, context) {
            if (values.rawValue && values.rawValue !== values.displayValue) {
                if (values.displayMarkdown) {
                    return `${renderCellWithHyperlink(cell, values.displayMarkdown, context)} [raw=${values.rawValue}]`;
                }
                return `[raw=${values.rawValue}]`;
            }
            return renderCellWithHyperlink(cell, values.displayMarkdown || values.rawValue, context);
        }
        function getCellMarkdownRenderer(outputMode) {
            if (outputMode === "raw") {
                return renderRawModeCell;
            }
            if (outputMode === "both") {
                return renderBothModeCell;
            }
            return renderDisplayModeCell;
        }
        function formatCellForMarkdown(cell, options, workbook = null, sheet = null) {
            if (!cell)
                return "";
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
        function isCellInAnyTable(row, col, tables) {
            return tables.some((table) => row >= table.startRow && row <= table.endRow && col >= table.startCol && col <= table.endCol);
        }
        function splitNarrativeRowSegments(cells, options, workbook = null, sheet = null) {
            const segments = [];
            let current = null;
            for (const cell of cells) {
                const value = formatCellForMarkdown(cell, options, workbook, sheet).trim();
                if (!value)
                    continue;
                if (!current || cell.col - current.lastCol > 4) {
                    current = {
                        startCol: cell.col,
                        values: [value],
                        lastCol: cell.col
                    };
                    segments.push(current);
                }
                else {
                    current.values.push(value);
                    current.lastCol = cell.col;
                }
            }
            return segments.map((segment) => ({
                startCol: segment.startCol,
                values: segment.values
            }));
        }
        function collectNarrativeCellsByRow(sheet, tables) {
            const rowMap = new Map();
            for (const cell of sheet.cells) {
                if (!cell.outputValue)
                    continue;
                if (isCellInAnyTable(cell.row, cell.col, tables))
                    continue;
                const entries = rowMap.get(cell.row) || [];
                entries.push(cell);
                rowMap.set(cell.row, entries);
            }
            return rowMap;
        }
        function createNarrativeItem(rowNumber, segment) {
            const text = segment.values.join(" ").trim();
            return {
                row: rowNumber,
                startCol: segment.startCol,
                text,
                cellValues: segment.values
            };
        }
        function createNarrativeItemsForRow(rowNumber, cells, options, workbook = null, sheet = null) {
            return splitNarrativeRowSegments(cells, options, workbook, sheet)
                .map((segment) => createNarrativeItem(rowNumber, segment))
                .filter((item) => !!item.text);
        }
        function buildNarrativeItems(workbook, sheet, tables, options = {}) {
            const rowMap = collectNarrativeCellsByRow(sheet, tables);
            const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
            const items = [];
            for (const rowNumber of rowNumbers) {
                const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
                items.push(...createNarrativeItemsForRow(rowNumber, cells, options, workbook, sheet));
            }
            return items;
        }
        function shouldStartNarrativeBlock(current, rowNumber, previousRow, startCol) {
            return !current || rowNumber - previousRow > 1 || Math.abs(startCol - current.startCol) > 3;
        }
        function isIsoDateToken(value) {
            return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
        }
        function isWeekdayToken(value) {
            const normalized = String(value || "").trim();
            return ["日", "月", "火", "水", "木", "金", "土", "日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"].includes(normalized);
        }
        function isCalendarDateItem(item) {
            if (!item)
                return false;
            const values = (item.cellValues || []).map((value) => String(value || "").trim()).filter(Boolean);
            return values.length >= 5 && values.every((value) => isIsoDateToken(value) || isWeekdayToken(value));
        }
        function blockHasCalendarDateItem(block) {
            return !!block?.items?.some((item) => isCalendarDateItem(item));
        }
        function shouldAppendToCalendarNarrativeBlock(current, item, previousRow) {
            if (!current || !blockHasCalendarDateItem(current)) {
                return false;
            }
            const rowGap = item.row - previousRow;
            if (rowGap < 0 || rowGap > 4) {
                return false;
            }
            const startColDelta = Math.abs(item.startCol - current.startCol);
            return startColDelta <= 24 || isCalendarDateItem(item);
        }
        function createNarrativeBlockFromItem(item) {
            return {
                startRow: item.row,
                startCol: item.startCol,
                endRow: item.row,
                lines: [item.text],
                items: [item]
            };
        }
        function appendNarrativeItem(block, item) {
            block.lines.push(item.text);
            block.endRow = item.row;
            block.items.push(item);
        }
        function extractNarrativeBlocks(workbook, sheet, tables, options = {}) {
            const items = buildNarrativeItems(workbook, sheet, tables, options);
            const blocks = [];
            let current = null;
            let previousRow = -100;
            for (const item of items) {
                if (shouldAppendToCalendarNarrativeBlock(current, item, previousRow)) {
                    appendNarrativeItem(current, item);
                }
                else if (shouldStartNarrativeBlock(current, item.row, previousRow, item.startCol)) {
                    current = createNarrativeBlockFromItem(item);
                    blocks.push(current);
                }
                else {
                    appendNarrativeItem(current, item);
                }
                previousRow = item.row;
            }
            return blocks;
        }
        function createNarrativeSectionAnchors(narrativeBlocks) {
            return narrativeBlocks.map((block) => ({
                startRow: block.startRow,
                startCol: block.startCol,
                endRow: block.endRow,
                endCol: Math.max(block.startCol, ...block.items.map((item) => item.startCol)),
                calendarLike: blockHasCalendarDateItem(block)
            }));
        }
        function createTableSectionAnchors(tables) {
            return tables.map((table) => ({
                startRow: table.startRow,
                startCol: table.startCol,
                endRow: table.endRow,
                endCol: table.endCol
            }));
        }
        function createPointSectionAnchor(address) {
            const anchor = deps.parseCellAddress(address);
            if (anchor.row > 0 && anchor.col > 0) {
                return { startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col };
            }
            return null;
        }
        function createImageSectionAnchors(sheet) {
            return sheet.images
                .map((image) => createPointSectionAnchor(image.anchor))
                .filter((anchor) => !!anchor);
        }
        function createChartSectionAnchors(charts) {
            return charts
                .map((chart) => createPointSectionAnchor(chart.anchor))
                .filter((anchor) => !!anchor);
        }
        function createSectionAnchors(sheet, tables, narrativeBlocks) {
            const charts = sheet.charts || [];
            return [
                ...createNarrativeSectionAnchors(narrativeBlocks),
                ...createTableSectionAnchors(tables),
                ...createImageSectionAnchors(sheet),
                ...createChartSectionAnchors(charts)
            ];
        }
        function sortSectionAnchors(anchors) {
            return anchors.sort((left, right) => {
                if (left.startRow !== right.startRow)
                    return left.startRow - right.startRow;
                return left.startCol - right.startCol;
            });
        }
        function createSectionBlockFromAnchor(anchor) {
            return {
                startRow: anchor.startRow,
                startCol: anchor.startCol,
                endRow: anchor.endRow,
                endCol: anchor.endCol
            };
        }
        function shouldStartNewSectionBlock(current, anchor, previousEndRow, previousAnchor) {
            const verticalGapThreshold = 4;
            const horizontalGapThreshold = 3;
            const gap = anchor.startRow - previousEndRow;
            if (!current) {
                return true;
            }
            if (gap > verticalGapThreshold) {
                const bothCalendarLike = !!(anchor.calendarLike && previousAnchor?.calendarLike);
                const horizontalGap = anchor.startCol - current.endCol;
                if (!(bothCalendarLike && gap <= 8 && horizontalGap <= 24)) {
                    return true;
                }
            }
            const overlapsCurrentRows = anchor.startRow <= current.endRow + 1;
            const horizontalGap = anchor.startCol - current.endCol;
            const bothCalendarLike = !!(anchor.calendarLike && previousAnchor?.calendarLike);
            if (bothCalendarLike && horizontalGap <= 24 && (overlapsCurrentRows || gap <= 8)) {
                return false;
            }
            if (overlapsCurrentRows && horizontalGap > horizontalGapThreshold) {
                return true;
            }
            return false;
        }
        function extendSectionBlock(section, anchor) {
            section.startRow = Math.min(section.startRow, anchor.startRow);
            section.startCol = Math.min(section.startCol, anchor.startCol);
            section.endRow = Math.max(section.endRow, anchor.endRow);
            section.endCol = Math.max(section.endCol, anchor.endCol);
        }
        function extractSectionBlocks(sheet, tables, narrativeBlocks) {
            const anchors = sortSectionAnchors(createSectionAnchors(sheet, tables, narrativeBlocks));
            if (anchors.length === 0) {
                return [];
            }
            const sections = [];
            let current = null;
            let previousEndRow = -100;
            let previousAnchor = null;
            for (const anchor of anchors) {
                if (shouldStartNewSectionBlock(current, anchor, previousEndRow, previousAnchor)) {
                    current = createSectionBlockFromAnchor(anchor);
                    sections.push(current);
                }
                else {
                    extendSectionBlock(current, anchor);
                }
                previousEndRow = Math.max(previousEndRow, anchor.endRow);
                previousAnchor = anchor;
            }
            return sections;
        }
        function createDefaultShapeSvgFilename(shapeIndex) {
            return `shape_${String(shapeIndex + 1).padStart(3, "0")}.svg`;
        }
        function createImageSectionEntry(image, index) {
            return [
                `### Image: ${String(index + 1).padStart(3, "0")} (${image.anchor})`,
                `- File: ${image.path}`,
                "",
                `![${image.filename}](${image.path})`
            ].join("\n");
        }
        function createChartSeriesLines(chart) {
            if (chart.series.length === 0) {
                return [];
            }
            const lines = ["- Series:"];
            for (const series of chart.series) {
                lines.push(`  - ${series.name}`);
                if (series.axis === "secondary")
                    lines.push("    - Axis: secondary");
                if (series.categoriesRef)
                    lines.push(`    - categories: ${series.categoriesRef}`);
                if (series.valuesRef)
                    lines.push(`    - values: ${series.valuesRef}`);
            }
            return lines;
        }
        function createChartSectionEntry(chart, index) {
            return [
                `### Chart: ${String(index + 1).padStart(3, "0")} (${chart.anchor})`,
                `- Title: ${chart.title || "(none)"}`,
                `- Type: ${chart.chartType}`,
                ...createChartSeriesLines(chart)
            ].join("\n");
        }
        function renderShapeDetails(shape, shapeIndex) {
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
        function createShapeBlockSummaryLine(shapeIndexes) {
            return shapeIndexes.map((shapeIndex) => `Shape ${String(shapeIndex + 1).padStart(3, "0")}`).join(", ");
        }
        function createShapeBlockEntry(block, blockIndex, shapes) {
            const shapeDetails = block.shapeIndexes
                .map((shapeIndex) => {
                const shape = shapes[shapeIndex];
                if (!shape)
                    return "";
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
        function collectUngroupedShapes(shapes, shapeBlocks) {
            const grouped = new Set(shapeBlocks.flatMap((block) => block.shapeIndexes));
            return shapes
                .map((shape, index) => ({ shape, index }))
                .filter(({ index }) => !grouped.has(index));
        }
        function renderImageSection(sheet) {
            return sheet.images.length > 0
                ? [
                    "",
                    ...sheet.images.map((image, index) => createImageSectionEntry(image, index))
                ].join("\n\n")
                : "";
        }
        function renderChartSection(charts) {
            return charts.length > 0
                ? [
                    "",
                    ...charts.map((chart, index) => createChartSectionEntry(chart, index))
                ].join("\n\n")
                : "";
        }
        function renderShapeSection(shapes, shapeBlocks, includeShapeDetails) {
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
        function createFormulaDiagnostics(sheet) {
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
        function createNarrativeSections(narrativeBlocks) {
            return narrativeBlocks.map((block) => ({
                sortRow: block.startRow,
                sortCol: block.startCol,
                markdown: `${deps.renderNarrativeBlock(block)}\n`,
                kind: "narrative",
                narrativeBlock: block
            }));
        }
        function createTableSections(workbook, sheet, tables, options, treatFirstRowAsHeader) {
            var _a;
            const sections = [];
            let tableCounter = 1;
            for (const table of tables) {
                const rows = deps.matrixFromCandidate(sheet, table, options, buildCellMap, (cell, tableOptions) => formatCellForMarkdown(cell, tableOptions, workbook, sheet));
                if (rows.length === 0 || ((_a = rows[0]) === null || _a === void 0 ? void 0 : _a.length) === 0)
                    continue;
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
        function sortContentSections(sections) {
            return sections.sort((left, right) => {
                if (left.sortRow !== right.sortRow)
                    return left.sortRow - right.sortRow;
                return left.sortCol - right.sortCol;
            });
        }
        function createFallbackSectionBlock() {
            return {
                startRow: -1,
                startCol: -1,
                endRow: Number.MAX_SAFE_INTEGER,
                endCol: Number.MAX_SAFE_INTEGER
            };
        }
        function isSectionInsideBlock(section, block) {
            return section.sortRow >= block.startRow
                && section.sortRow <= block.endRow
                && section.sortCol >= block.startCol
                && section.sortCol <= block.endCol;
        }
        function createGroupedSections(sectionBlocks, sections) {
            const blocks = sectionBlocks.length > 0
                ? sectionBlocks
                : [createFallbackSectionBlock()];
            return blocks.map((block) => ({
                block,
                entries: sections.filter((section) => isSectionInsideBlock(section, block))
            })).filter((group) => group.entries.length > 0);
        }
        function renderGroupedSectionEntries(entries) {
            return createCalendarAwareSectionEntries(entries).map((section) => section.markdown.trimEnd()).join("\n\n").trim();
        }
        function isIsoDateToken(value) {
            return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
        }
        function isWeekdayToken(value) {
            const normalized = String(value || "").trim();
            return ["日", "月", "火", "水", "木", "金", "土", "日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"].includes(normalized);
        }
        function getNarrativeValues(section) {
            return section.narrativeBlock?.items.flatMap((item) => item.cellValues || []).map((value) => String(value || "").trim()).filter(Boolean) || [];
        }
        function isCalendarBodySection(section) {
            if (section.kind !== "narrative" || !section.narrativeBlock)
                return false;
            return section.narrativeBlock.items.some((item) => {
                const values = (item.cellValues || []).map((value) => String(value || "").trim()).filter(Boolean);
                return values.length >= 5 && values.every((value) => isIsoDateToken(value) || isWeekdayToken(value));
            });
        }
        function isCalendarHeaderSection(section) {
            if (section.kind !== "narrative" || !section.narrativeBlock)
                return false;
            const values = getNarrativeValues(section);
            if (values.length === 0 || values.length > 10)
                return false;
            const allWeekdays = values.length >= 5 && values.every((value) => isWeekdayToken(value));
            const hasMonthTitle = values.some((value) => /^\d{4}年\d{1,2}月$/.test(value));
            const hasPlannerLabel = values.includes("目標と優先事項") || values.includes("その他");
            return allWeekdays || hasMonthTitle || hasPlannerLabel;
        }
        function createSyntheticSection(markdown, sortRow, sortCol) {
            return {
                sortRow,
                sortCol,
                markdown,
                kind: "narrative"
            };
        }
        function createCalendarAwareSectionEntries(entries) {
            const mainCalendarEntries = entries.filter((entry) => isCalendarBodySection(entry))
                .sort((left, right) => {
                if (left.sortCol !== right.sortCol)
                    return left.sortCol - right.sortCol;
                return left.sortRow - right.sortRow;
            });
            if (mainCalendarEntries.length === 0) {
                return entries;
            }
            const main = mainCalendarEntries[0];
            const headerEntries = entries.filter((entry) => (entry !== main
                && isCalendarHeaderSection(entry)
                && entry.sortRow <= main.sortRow + 1));
            const sidebarEntries = entries.filter((entry) => (entry !== main
                && !headerEntries.includes(entry)
                && isCalendarBodySection(entry)
                && entry.sortCol >= main.sortCol + 10));
            const remainingEntries = entries.filter((entry) => (entry !== main
                && !headerEntries.includes(entry)
                && !sidebarEntries.includes(entry)));
            if (headerEntries.length === 0 && sidebarEntries.length === 0) {
                return entries;
            }
            const reordered = [];
            reordered.push(...headerEntries);
            reordered.push(main);
            if (sidebarEntries.length > 0) {
                const firstSidebar = sidebarEntries[0];
                reordered.push(createSyntheticSection("### Sidebar\n", firstSidebar.sortRow - 0.1, firstSidebar.sortCol));
                reordered.push(...sidebarEntries);
            }
            reordered.push(...remainingEntries);
            return reordered;
        }
        function renderGroupedSectionBody(groupedSections) {
            return groupedSections
                .map((group) => renderGroupedSectionEntries(group.entries))
                .filter(Boolean)
                .join("\n\n---\n\n")
                .trim();
        }
        function collectSheetRenderState(workbook, sheet, options = {}) {
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
        function createSheetMarkdownText(workbook, sheet, state) {
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
        function createSheetSummary(sheet, state) {
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
        function convertSheetToMarkdown(workbook, sheet, options = {}) {
            const state = collectSheetRenderState(workbook, sheet, options);
            const fileName = deps.createOutputFileName(workbook.name, sheet.index, sheet.name, state.resolvedOptions.outputMode, state.resolvedOptions.formattingMode);
            return {
                fileName,
                sheetName: sheet.name,
                markdown: createSheetMarkdownText(workbook, sheet, state),
                summary: createSheetSummary(sheet, state)
            };
        }
        function convertWorkbookToMarkdownFiles(workbook, options = {}) {
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
            createCalendarAwareSectionEntries,
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
