/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();

  type MarkdownOptions = {
    treatFirstRowAsHeader?: boolean;
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
    includeShapeDetails?: boolean;
    outputMode?: "display" | "raw" | "both" | string;
    formattingMode?: "plain" | "github" | string;
    tableDetectionMode?: "balanced" | "border" | "planner-aware" | string;
  };

  type ResolvedMarkdownOptions = {
    treatFirstRowAsHeader: boolean;
    trimText: boolean;
    removeEmptyRows: boolean;
    removeEmptyColumns: boolean;
    includeShapeDetails: boolean;
    outputMode: "display" | "raw" | "both";
    formattingMode: "plain" | "github";
    tableDetectionMode: "balanced" | "border" | "planner-aware";
  };

  const OUTPUT_MODES = ["display", "raw", "both"] as const;
  const FORMATTING_MODES = ["plain", "github"] as const;
  const TABLE_DETECTION_MODES = ["balanced", "border", "planner-aware"] as const;
  const TABLE_DETECTION_MODE_ALIASES: Record<string, "balanced" | "border" | "planner-aware"> = {
    "border-priority": "border"
  };

  function normalizeEnum<T extends string>(
    value: string | null | undefined,
    allowedValues: readonly T[],
    fallback: T,
    aliases: Record<string, T> = {}
  ): T {
    const normalizedInput = String(value || "").trim().toLowerCase();
    if (!normalizedInput) {
      return fallback;
    }
    const normalizedValue = aliases[normalizedInput] || normalizedInput;
    return (allowedValues as readonly string[]).includes(normalizedValue)
      ? normalizedValue as T
      : fallback;
  }

  function resolveBoolean(value: boolean | null | undefined, fallback: boolean): boolean {
    return value === undefined || value === null ? fallback : value !== false;
  }

  function normalizeOutputMode(value?: string | null): "display" | "raw" | "both" {
    return normalizeEnum(value, OUTPUT_MODES, "display");
  }

  function normalizeFormattingMode(value?: string | null): "plain" | "github" {
    return normalizeEnum(value, FORMATTING_MODES, "plain");
  }

  function normalizeTableDetectionMode(value?: string | null): "balanced" | "border" | "planner-aware" {
    return normalizeEnum(value, TABLE_DETECTION_MODES, "balanced", TABLE_DETECTION_MODE_ALIASES);
  }

  function resolveMarkdownOptions(options: MarkdownOptions = {}): ResolvedMarkdownOptions {
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
