/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const OUTPUT_MODES = ["display", "raw", "both"];
    const FORMATTING_MODES = ["plain", "github"];
    const TABLE_DETECTION_MODES = ["balanced", "border"];
    const TABLE_DETECTION_MODE_ALIASES = {
        "border-priority": "border"
    };
    function normalizeEnum(value, allowedValues, fallback, aliases = {}) {
        const normalizedInput = String(value || "").trim().toLowerCase();
        if (!normalizedInput) {
            return fallback;
        }
        const normalizedValue = aliases[normalizedInput] || normalizedInput;
        return allowedValues.includes(normalizedValue)
            ? normalizedValue
            : fallback;
    }
    function resolveBoolean(value, fallback) {
        return value === undefined || value === null ? fallback : value !== false;
    }
    function normalizeOutputMode(value) {
        return normalizeEnum(value, OUTPUT_MODES, "display");
    }
    function normalizeFormattingMode(value) {
        return normalizeEnum(value, FORMATTING_MODES, "plain");
    }
    function normalizeTableDetectionMode(value) {
        return normalizeEnum(value, TABLE_DETECTION_MODES, "balanced", TABLE_DETECTION_MODE_ALIASES);
    }
    function resolveMarkdownOptions(options = {}) {
        return {
            treatFirstRowAsHeader: resolveBoolean(options.treatFirstRowAsHeader, true),
            trimText: resolveBoolean(options.trimText, true),
            removeEmptyRows: resolveBoolean(options.removeEmptyRows, true),
            removeEmptyColumns: resolveBoolean(options.removeEmptyColumns, true),
            includeShapeDetails: resolveBoolean(options.includeShapeDetails, true),
            outputMode: normalizeOutputMode(options.outputMode),
            formattingMode: normalizeFormattingMode(options.formattingMode),
            tableDetectionMode: normalizeTableDetectionMode(options.tableDetectionMode)
        };
    }
    moduleRegistry.registerModule("markdownOptions", {
        OUTPUT_MODES,
        FORMATTING_MODES,
        TABLE_DETECTION_MODES,
        TABLE_DETECTION_MODE_ALIASES,
        normalizeOutputMode,
        normalizeFormattingMode,
        normalizeTableDetectionMode,
        resolveMarkdownOptions
    });
})();
