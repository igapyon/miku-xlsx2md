/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const EMPTY_SUMMARY_HTML = '<div class="md-summary-empty">No conversion yet.</div>';
    const EMPTY_TABLE_SCORE_HTML = '<div class="md-summary-empty">No table candidates found.</div>';
    const EMPTY_FORMULA_HTML = '<div class="md-summary-empty">No formula cells found.</div>';
    const TABLE_SCORE_LABELS = ["strong", "candidate", "unknown"];
    const FORMULA_STATUS_LABELS = ["resolved", "fallback", "unsupported", "unknown"];
    const FORMULA_SOURCE_LABELS = ["cached", "ast", "legacy", "formula", "external", "unknown"];
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const textEncoding = requireXlsx2mdTextEncoding();
    const markdownOptions = requireXlsx2mdMarkdownOptions();
    const xlsx2md = moduleRegistry.getModule("xlsx2md");
    if (!xlsx2md) {
        throw new Error("xlsx2md core module is not loaded");
    }
    let currentWorkbook = null;
    let currentFiles = [];
    let currentWorkbookBytes = null;
    let currentWorkbookName = "";
    let currentParsedIncludeShapeDetails = null;
    function getElement(id) {
        const element = document.getElementById(id);
        if (!element) {
            throw new Error(`Element not found: ${id}`);
        }
        return element;
    }
    function getSwitchValue(id) {
        const element = getElement(id);
        return !!element.checked;
    }
    function getSelectValue(id, fallback) {
        var _a;
        const element = getElement(id);
        if (typeof element.getValue === "function") {
            return element.getValue() || fallback;
        }
        return ((_a = document.getElementById(id)) === null || _a === void 0 ? void 0 : _a.value) || fallback;
    }
    function isEncodingAvailable(encoding) {
        return textEncoding.isEncodingAvailable(encoding);
    }
    function getOptions() {
        return markdownOptions.resolveMarkdownOptions({
            treatFirstRowAsHeader: getSwitchValue("headerRowEnabled"),
            trimText: getSwitchValue("trimTextEnabled"),
            removeEmptyRows: getSwitchValue("removeEmptyRowsEnabled"),
            removeEmptyColumns: getSwitchValue("removeEmptyColumnsEnabled"),
            includeShapeDetails: getSwitchValue("includeShapeDetailsEnabled"),
            outputMode: getSelectValue("outputModeSelect", "display"),
            formattingMode: getSelectValue("formattingModeSelect", "plain"),
            tableDetectionMode: getSelectValue("tableDetectionModeSelect", "balanced")
        });
    }
    function getEncodingOptions() {
        const encoding = getSelectValue("encodingSelect", "utf-8");
        const bom = getSelectValue("bomSelect", "off");
        return {
            encoding: (encoding === "shift_jis" ||
                encoding === "utf-16le" ||
                encoding === "utf-16be" ||
                encoding === "utf-32le" ||
                encoding === "utf-32be") ? encoding : "utf-8",
            bom: bom === "on" ? "on" : "off"
        };
    }
    function getSelectedOutputMode() {
        return getOptions().outputMode;
    }
    function showToast(message) {
        const toast = document.getElementById("toast");
        if (toast && typeof toast.show === "function") {
            toast.show(message, 2200);
        }
    }
    function setSummaryHtml(html) {
        getElement("analysisSummary").innerHTML = html;
    }
    function setSummaryText(message) {
        setSummaryHtml(`<div class="md-summary-empty">${escapeHtml(message)}</div>`);
    }
    function setScoreSummaryHtml(html) {
        getElement("scoreSummary").innerHTML = html;
    }
    function setFormulaSummaryHtml(html) {
        getElement("formulaSummary").innerHTML = html;
    }
    function escapeHtml(value) {
        return String(value || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }
    function countByLabel(items, getLabel, labels) {
        const counts = Object.fromEntries(labels.map((label) => [label, 0]));
        items.forEach((item) => {
            counts[getLabel(item)] += 1;
        });
        return counts;
    }
    function formatNonZeroCounts(counts, labels) {
        return labels
            .filter((label) => counts[label] > 0)
            .map((label) => `${label} ${counts[label]}`)
            .join(" / ");
    }
    function sumFileMetric(files, getValue) {
        return files.reduce((sum, file) => sum + getValue(file), 0);
    }
    function resetRenderedWorkbookState(summaryMessage) {
        setSummaryText(summaryMessage);
        setScoreSummaryHtml(EMPTY_SUMMARY_HTML);
        setFormulaSummaryHtml(EMPTY_SUMMARY_HTML);
        setPreviewMarkdown("");
        getElement("downloadBtn").disabled = true;
        getElement("exportZipBtn").disabled = true;
    }
    function renderScoreSummary(files) {
        const sheetsWithScores = files.filter((file) => file.summary.tableScores.length > 0);
        if (sheetsWithScores.length === 0) {
            return EMPTY_TABLE_SCORE_HTML;
        }
        const allScores = sheetsWithScores.flatMap((file) => file.summary.tableScores);
        const totalCounts = countByLabel(allScores, (detail) => getTableScoreLabel(detail.score), TABLE_SCORE_LABELS);
        return `<div class="md-summary-overview">Total ${allScores.length} / strong ${totalCounts.strong} / candidate ${totalCounts.candidate}</div>${sheetsWithScores.map((file) => {
            const items = [...file.summary.tableScores].sort((left, right) => {
                if (right.score !== left.score) {
                    return right.score - left.score;
                }
                return left.range.localeCompare(right.range);
            }).map((detail) => (`<div class="md-summary-item"><div class="md-summary-item-head"><span class="md-summary-item-title">${escapeHtml(detail.range)}</span><span class="md-summary-item-status md-summary-item-status--${escapeHtml(getTableScoreLabel(detail.score))}">${escapeHtml(getTableScoreText(detail.score))}</span></div><div class="md-summary-item-meta">Score ${detail.score}</div><div class="md-summary-item-body">${escapeHtml(detail.reasons.join(" / "))}</div></div>`)).join("");
            return `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">${escapeHtml(file.sheetName)}</h3><span class="md-summary-group-count">${file.summary.tableScores.length}</span></div><div class="md-summary-group-meta">${escapeHtml(renderTableScoreCounts(file))}</div>${items}</section>`;
        }).join("")}`;
    }
    function getTableScoreLabel(score) {
        if (score >= 7)
            return "strong";
        if (score >= 4)
            return "candidate";
        return "unknown";
    }
    function getTableScoreText(score) {
        if (score >= 7)
            return "strong";
        if (score >= 4)
            return "candidate";
        return "unknown";
    }
    function renderTableScoreCounts(file) {
        return formatNonZeroCounts(countByLabel(file.summary.tableScores, (detail) => getTableScoreLabel(detail.score), TABLE_SCORE_LABELS), TABLE_SCORE_LABELS);
    }
    function getFormulaStatusLabel(status) {
        if (status === "resolved")
            return "resolved";
        if (status === "fallback_formula")
            return "fallback";
        if (status === "unsupported_external")
            return "unsupported";
        return "unknown";
    }
    function renderFormulaStatusCounts(file) {
        return formatNonZeroCounts(countByLabel(file.summary.formulaDiagnostics, (diagnostic) => getFormulaStatusLabel(diagnostic.status), FORMULA_STATUS_LABELS), FORMULA_STATUS_LABELS);
    }
    function getFormulaStatusPriority(status) {
        const label = getFormulaStatusLabel(status);
        if (label === "unsupported")
            return 0;
        if (label === "fallback")
            return 1;
        if (label === "unknown")
            return 2;
        return 3;
    }
    function getFormulaSourceLabel(source) {
        if (source === "cached_value")
            return "cached";
        if (source === "ast_evaluator")
            return "ast";
        if (source === "legacy_resolver")
            return "legacy";
        if (source === "formula_text")
            return "formula";
        if (source === "external_unsupported")
            return "external";
        return "unknown";
    }
    function renderFormulaSourceCounts(file) {
        return formatNonZeroCounts(countByLabel(file.summary.formulaDiagnostics, (diagnostic) => getFormulaSourceLabel(diagnostic.source), FORMULA_SOURCE_LABELS), FORMULA_SOURCE_LABELS);
    }
    function renderFormulaSummary(files) {
        const sheetsWithDiagnostics = files.filter((file) => file.summary.formulaDiagnostics.length > 0);
        if (sheetsWithDiagnostics.length === 0) {
            return EMPTY_FORMULA_HTML;
        }
        const diagnostics = sheetsWithDiagnostics.flatMap((file) => file.summary.formulaDiagnostics);
        const totalStatuses = countByLabel(diagnostics, (diagnostic) => getFormulaStatusLabel(diagnostic.status), FORMULA_STATUS_LABELS);
        const totalSources = countByLabel(diagnostics, (diagnostic) => getFormulaSourceLabel(diagnostic.source), FORMULA_SOURCE_LABELS);
        return `<div class="md-summary-overview">Total ${diagnostics.length} / cached ${totalSources.cached} / ast ${totalSources.ast} / legacy ${totalSources.legacy} / formula ${totalSources.formula} / resolved ${totalStatuses.resolved} / fallback ${totalStatuses.fallback} / unsupported ${totalStatuses.unsupported}</div>${sheetsWithDiagnostics.map((file) => {
            const items = [...file.summary.formulaDiagnostics].sort((left, right) => {
                const priorityDiff = getFormulaStatusPriority(left.status) - getFormulaStatusPriority(right.status);
                if (priorityDiff !== 0) {
                    return priorityDiff;
                }
                return left.address.localeCompare(right.address);
            }).map((diagnostic) => (`<div class="md-summary-item"><div class="md-summary-item-head"><span class="md-summary-item-title">${escapeHtml(diagnostic.address)}</span><span class="md-summary-item-badges"><span class="md-summary-item-status md-summary-item-status--source-${escapeHtml(getFormulaSourceLabel(diagnostic.source))}">${escapeHtml(getFormulaSourceLabel(diagnostic.source))}</span><span class="md-summary-item-status md-summary-item-status--${escapeHtml(getFormulaStatusLabel(diagnostic.status))}">${escapeHtml(getFormulaStatusLabel(diagnostic.status))}</span></span></div><div class="md-summary-item-body">${escapeHtml(`${diagnostic.formulaText} => ${diagnostic.outputValue}`)}</div></div>`)).join("");
            return `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">${escapeHtml(file.sheetName)}</h3><span class="md-summary-group-count">${file.summary.formulaDiagnostics.length}</span></div><div class="md-summary-group-meta">${escapeHtml(renderFormulaSourceCounts(file))}</div><div class="md-summary-group-meta">${escapeHtml(renderFormulaStatusCounts(file))}</div>${items}</section>`;
        }).join("")}`;
    }
    function renderAnalysisSummary(files, workbookName) {
        var _a, _b, _c;
        if (files.length === 0) {
            return EMPTY_SUMMARY_HTML;
        }
        const totalTables = sumFileMetric(files, (file) => file.summary.tables);
        const totalNarratives = sumFileMetric(files, (file) => file.summary.narrativeBlocks);
        const totalMerges = sumFileMetric(files, (file) => file.summary.merges);
        const totalImages = sumFileMetric(files, (file) => file.summary.images);
        const totalCells = sumFileMetric(files, (file) => file.summary.cells);
        const totalFormulas = sumFileMetric(files, (file) => file.summary.formulaDiagnostics.length);
        const outputMode = ((_a = files[0]) === null || _a === void 0 ? void 0 : _a.summary.outputMode) || "display";
        const formattingMode = ((_b = files[0]) === null || _b === void 0 ? void 0 : _b.summary.formattingMode) || "plain";
        const tableDetectionMode = ((_c = files[0]) === null || _c === void 0 ? void 0 : _c.summary.tableDetectionMode) || "balanced";
        const overview = `<div class="md-summary-overview">Workbook ${escapeHtml(workbookName)} / ${files.length} sheet(s) / value mode ${escapeHtml(outputMode)} / formatting ${escapeHtml(formattingMode)} / table detection ${escapeHtml(tableDetectionMode)}</div>`;
        const items = files.map((file) => (`<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">${escapeHtml(file.sheetName)}</h3><span class="md-summary-group-count">${file.summary.cells} cells</span></div><div class="md-summary-group-meta">tables ${file.summary.tables} / narrative ${file.summary.narrativeBlocks} / merges ${file.summary.merges} / images ${file.summary.images} / formulas ${file.summary.formulaDiagnostics.length}</div></section>`)).join("");
        const totals = `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">Total</h3><span class="md-summary-group-count">${files.length} sheets</span></div><div class="md-summary-group-meta">tables ${totalTables} / narrative ${totalNarratives} / merges ${totalMerges} / images ${totalImages} / formulas ${totalFormulas} / analyzed cells ${totalCells}</div></section>`;
        return `${overview}${totals}${items}`;
    }
    function updateOutputModeNotice(mode) {
        const notice = getElement("outputModeNotice");
        if (mode === "raw") {
            notice.textContent = "`raw` outputs internal values instead of Excel's displayed values.";
            return;
        }
        if (mode === "both") {
            notice.textContent = "`both` outputs displayed values plus supplemental `[raw=...]` data.";
            return;
        }
        notice.textContent = "`display` outputs values close to what Excel shows.";
    }
    function updateFormattingModeNotice(mode) {
        const notice = getElement("formattingModeNotice");
        if (mode === "github") {
            notice.textContent = "`github` preserves supported Excel emphasis as GitHub-compatible Markdown: bold, italic, strike, underline, and in-cell line breaks as `<br>`.";
            return;
        }
        notice.textContent = "`plain` strips Excel text emphasis and outputs plain Markdown text.";
    }
    function updateTableDetectionModeNotice(mode) {
        const notice = getElement("tableDetectionModeNotice");
        if (mode === "border") {
            notice.textContent = "`border` detects tables from bordered regions and suppresses borderless fallback detection.";
            return;
        }
        notice.textContent = "`balanced` uses both bordered candidates and value-density fallback detection.";
    }
    function updateEncodingNotice(encoding) {
        const notice = getElement("encodingNotice");
        if (encoding === "shift_jis") {
            notice.textContent = isEncodingAvailable("shift_jis")
                ? "`shift_jis` save is available in this runtime, including the Node CLI path."
                : "`shift_jis` save is not available in this browser runtime. Use the Node CLI for Shift_JIS output.";
            return;
        }
        if (encoding === "utf-16le" || encoding === "utf-16be" || encoding === "utf-32le" || encoding === "utf-32be") {
            notice.textContent = `\`${encoding}\` writes Unicode text in the selected endian form instead of UTF-8.`;
            return;
        }
        notice.textContent = "`utf-8` is the default Markdown encoding.";
    }
    function updateBomNotice(options) {
        const notice = getElement("bomNotice");
        if (options.encoding === "shift_jis") {
            notice.textContent = "`shift_jis` does not support BOM output.";
            return;
        }
        if (options.bom === "on") {
            notice.textContent = "BOM will be written at the start of the saved Markdown bytes.";
            return;
        }
        notice.textContent = "BOM is disabled for saved Markdown bytes.";
    }
    function syncEncodingControls() {
        const encodingSelect = getElement("encodingSelect");
        const shiftJisOption = Array.from(encodingSelect.options).find((option) => option.value === "shift_jis") || null;
        const shiftJisAvailable = isEncodingAvailable("shift_jis");
        if (shiftJisOption) {
            shiftJisOption.disabled = !shiftJisAvailable;
            shiftJisOption.text = shiftJisAvailable ? "Shift_JIS" : "Shift_JIS (CLI only)";
        }
        if (!shiftJisAvailable && encodingSelect.value === "shift_jis") {
            encodingSelect.value = "utf-8";
        }
        const options = getEncodingOptions();
        const bomSelect = getElement("bomSelect");
        if (options.encoding === "shift_jis") {
            bomSelect.value = "off";
            bomSelect.disabled = true;
        }
        else {
            bomSelect.disabled = false;
        }
        updateEncodingNotice(options.encoding);
        updateBomNotice(getEncodingOptions());
    }
    function updatePreviewModeBanner(mode) {
        const banner = getElement("previewModeBanner");
        if (mode === "raw") {
            banner.hidden = false;
            banner.textContent = "`raw` mode is active. Markdown will show internal values instead of Excel's displayed values.";
            return;
        }
        if (mode === "both") {
            banner.hidden = false;
            banner.textContent = "`both` mode is active. Markdown will include displayed values plus `[raw=...]` annotations.";
            return;
        }
        banner.hidden = true;
        banner.textContent = "";
    }
    function setPreviewMarkdown(markdown) {
        const preview = getElement("markdownPreview");
        if (typeof preview.setText === "function") {
            preview.setText(markdown);
            return;
        }
        getElement("markdownOutput").textContent = markdown;
    }
    function clearError() {
        const errorAlert = getElement("errorAlert");
        if (typeof errorAlert.clear === "function") {
            errorAlert.clear();
        }
        else {
            errorAlert.removeAttribute("active");
            errorAlert.textContent = "";
        }
    }
    function showError(message) {
        const errorAlert = getElement("errorAlert");
        if (typeof errorAlert.show === "function") {
            errorAlert.show(message);
        }
        else {
            errorAlert.textContent = message;
            errorAlert.setAttribute("active", "");
        }
    }
    function setLoading(active, message) {
        const overlay = getElement("loadingOverlay");
        if (active) {
            if (message) {
                overlay.setAttribute("text", message);
            }
            if (typeof overlay.show === "function") {
                overlay.show(message || "Processing");
            }
            else {
                overlay.setAttribute("active", "");
            }
            return;
        }
        if (typeof overlay.hide === "function") {
            overlay.hide();
        }
        else {
            overlay.removeAttribute("active");
        }
    }
    function renderCurrentSelection() {
        var _a;
        if (!currentFiles.length) {
            resetRenderedWorkbookState("No conversion yet.");
            updatePreviewModeBanner(getSelectedOutputMode());
            return;
        }
        if (!currentWorkbook) {
            showError("Workbook context is missing.");
            return;
        }
        const combinedMarkdown = xlsx2md.createCombinedMarkdownExportFile(currentWorkbook, currentFiles).content;
        const outputMode = ((_a = currentFiles[0]) === null || _a === void 0 ? void 0 : _a.summary.outputMode) || "display";
        updatePreviewModeBanner(outputMode);
        setSummaryHtml(renderAnalysisSummary(currentFiles, (currentWorkbook === null || currentWorkbook === void 0 ? void 0 : currentWorkbook.name) || "workbook.xlsx"));
        setScoreSummaryHtml(renderScoreSummary(currentFiles));
        setFormulaSummaryHtml(renderFormulaSummary(currentFiles));
        setPreviewMarkdown(combinedMarkdown);
        getElement("downloadBtn").disabled = false;
        getElement("exportZipBtn").disabled = false;
    }
    function getSelectedFileForDownload() {
        if (!currentFiles.length)
            return null;
        if (!currentWorkbook)
            return null;
        return xlsx2md.createCombinedMarkdownExportPayload(currentWorkbook, currentFiles, getEncodingOptions());
    }
    function downloadCurrentMarkdown() {
        try {
            const payload = getSelectedFileForDownload();
            if (!payload) {
                showError("No Markdown is available to save.");
                return;
            }
            const blob = new Blob([payload.data], { type: payload.mimeType });
            const objectUrl = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = objectUrl;
            link.download = payload.fileName;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
            showToast("Saved Markdown.");
        }
        catch (error) {
            showError(error instanceof Error ? error.message : "Failed to save Markdown.");
        }
    }
    function downloadExportZip() {
        try {
            if (!currentWorkbook || currentFiles.length === 0) {
                showError("Generate Markdown first.");
                return;
            }
            const zipBytes = xlsx2md.createWorkbookExportArchive(currentWorkbook, currentFiles, getEncodingOptions());
            const blob = new Blob([zipBytes], { type: "application/zip" });
            const objectUrl = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = objectUrl;
            link.download = `${currentWorkbook.name.replace(/\.xlsx$/i, "")}.zip`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
            showToast("Saved ZIP archive.");
        }
        catch (error) {
            showError(error instanceof Error ? error.message : "Failed to save ZIP archive.");
        }
    }
    async function ensureWorkbookParsedForCurrentOptions() {
        if (!currentWorkbookBytes || !currentWorkbookName) {
            return currentWorkbook !== null;
        }
        const includeShapeDetails = getOptions().includeShapeDetails;
        if (currentWorkbook && currentParsedIncludeShapeDetails === includeShapeDetails) {
            return true;
        }
        setLoading(true, "Analyzing xlsx");
        try {
            currentWorkbook = await xlsx2md.parseWorkbook(currentWorkbookBytes, currentWorkbookName, { includeShapeDetails });
            currentParsedIncludeShapeDetails = includeShapeDetails;
            return true;
        }
        catch (error) {
            showError(error instanceof Error ? error.message : "Failed to load the xlsx file.");
            return false;
        }
        finally {
            setLoading(false);
        }
    }
    async function convertCurrentWorkbook(showSuccessToast = true) {
        clearError();
        if (!currentWorkbook && !currentWorkbookBytes) {
            showError("Load an xlsx file first.");
            return;
        }
        try {
            const ready = await ensureWorkbookParsedForCurrentOptions();
            if (!ready || !currentWorkbook) {
                return;
            }
            currentFiles = xlsx2md.convertWorkbookToMarkdownFiles(currentWorkbook, getOptions());
            renderCurrentSelection();
            if (showSuccessToast) {
                showToast("Generated Markdown.");
            }
        }
        catch (error) {
            showError(error instanceof Error ? error.message : "Failed to generate Markdown.");
        }
    }
    async function loadWorkbookFromFile(file) {
        clearError();
        setLoading(true, "Loading xlsx");
        try {
            const arrayBuffer = await file.arrayBuffer();
            currentWorkbookBytes = arrayBuffer;
            currentWorkbookName = file.name;
            currentWorkbook = await xlsx2md.parseWorkbook(arrayBuffer, file.name, {
                includeShapeDetails: getOptions().includeShapeDetails
            });
            currentParsedIncludeShapeDetails = getOptions().includeShapeDetails;
            currentFiles = [];
            await convertCurrentWorkbook(false);
            showToast("Loaded xlsx and generated Markdown.");
        }
        catch (error) {
            currentWorkbook = null;
            currentFiles = [];
            currentWorkbookBytes = null;
            currentWorkbookName = "";
            currentParsedIncludeShapeDetails = null;
            resetRenderedWorkbookState("Failed to load the workbook.");
            showError(error instanceof Error ? error.message : "Failed to load the xlsx file.");
        }
        finally {
            setLoading(false);
        }
    }
    function bindFileInput() {
        const fileInput = getElement("xlsxFileInput");
        fileInput.addEventListener("change", async () => {
            var _a;
            const file = (_a = fileInput.files) === null || _a === void 0 ? void 0 : _a[0];
            if (!file)
                return;
            await loadWorkbookFromFile(file);
        });
    }
    function bindActions() {
        getElement("convertBtn").addEventListener("click", () => {
            void convertCurrentWorkbook(true);
        });
        getElement("downloadBtn").addEventListener("click", () => {
            downloadCurrentMarkdown();
        });
        getElement("exportZipBtn").addEventListener("click", () => {
            downloadExportZip();
        });
        getElement("outputModeSelect").addEventListener("change", () => {
            const mode = getSelectedOutputMode();
            updateOutputModeNotice(mode);
            if (!currentFiles.length) {
                updatePreviewModeBanner(mode);
            }
        });
        getElement("formattingModeSelect").addEventListener("change", () => {
            updateFormattingModeNotice(getOptions().formattingMode);
        });
        getElement("tableDetectionModeSelect").addEventListener("change", () => {
            updateTableDetectionModeNotice(getOptions().tableDetectionMode);
        });
        getElement("encodingSelect").addEventListener("change", () => {
            syncEncodingControls();
        });
        getElement("bomSelect").addEventListener("change", () => {
            updateBomNotice(getEncodingOptions());
        });
    }
    function initialize() {
        clearError();
        resetRenderedWorkbookState("No conversion yet.");
        updateOutputModeNotice(getSelectedOutputMode());
        updateFormattingModeNotice(getOptions().formattingMode);
        updateTableDetectionModeNotice(getOptions().tableDetectionMode);
        syncEncodingControls();
        updatePreviewModeBanner(getSelectedOutputMode());
        bindFileInput();
        bindActions();
    }
    document.addEventListener("DOMContentLoaded", initialize);
})();
