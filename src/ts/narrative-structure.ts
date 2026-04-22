/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();

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

  function normalizeNarrativeText(text: string): string {
    return markdownNormalizeHelper.normalizeMarkdownText(text);
  }

  function formatNarrativeHeading(text: string): string {
    return `### ${markdownNormalizeHelper.normalizeMarkdownHeadingText(text)}`;
  }

  function formatNarrativeBullet(text: string): string {
    return `- ${markdownNormalizeHelper.normalizeMarkdownListItemText(text)}`;
  }

  function isIndentedChildItem(parent: NarrativeItem | null | undefined, child: NarrativeItem | null | undefined): boolean {
    return !!(parent && child && child.startCol > parent.startCol);
  }

  function isWeekdayToken(value: string): boolean {
    const normalized = String(value || "").trim();
    return ["日", "月", "火", "水", "木", "金", "土", "日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"].includes(normalized);
  }

  function isIsoDateToken(value: string): boolean {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
  }

  function isCalendarLikeItem(item: NarrativeItem | null | undefined): boolean {
    if (!item || !Array.isArray(item.cellValues)) {
      return false;
    }
    const values = item.cellValues.map((value) => String(value || "").trim()).filter(Boolean);
    if (values.length < 5) {
      return false;
    }
    const weekdayCount = values.filter(isWeekdayToken).length;
    const dateCount = values.filter(isIsoDateToken).length;
    return weekdayCount >= 5 || dateCount >= 5 || values.length >= 7;
  }

  function isCalendarLikeNarrativeBlock(block: NarrativeBlock): boolean {
    if (!block.items || block.items.length < 2) {
      return false;
    }
    const calendarLikeItems = block.items.filter((item) => isCalendarLikeItem(item));
    return calendarLikeItems.length >= 2;
  }

  function renderCalendarLikeItem(item: NarrativeItem): string {
    const values = item.cellValues.map((value) => String(value || "").trim()).filter(Boolean);
    if (values.length === 0) {
      return normalizeNarrativeText(item.text);
    }
    if (values.every(isWeekdayToken)) {
      return formatNarrativeHeading(values.join(" "));
    }
    if (values.every((value) => isIsoDateToken(value) || isWeekdayToken(value))) {
      return values.join(" | ");
    }
    return values.join(" | ");
  }

  function renderCalendarLikeNarrativeBlock(block: NarrativeBlock): string {
    return block.items
      .map((item) => {
        const values = item.cellValues.map((value) => String(value || "").trim()).filter(Boolean);
        if (isCalendarLikeItem(item) || values.length >= 2) {
          return renderCalendarLikeItem(item);
        }
        return normalizeNarrativeText(item.text);
      })
      .join("\n\n");
  }

  function renderNarrativeBlock(block: NarrativeBlock): string {
    if (!block.items || block.items.length === 0) {
      return block.lines.map((line) => normalizeNarrativeText(line)).join("\n");
    }
    if (isCalendarLikeNarrativeBlock(block)) {
      return renderCalendarLikeNarrativeBlock(block);
    }
    const parts: string[] = [];
    let index = 0;
    while (index < block.items.length) {
      const current = block.items[index];
      const next = block.items[index + 1];
      if (isIndentedChildItem(current, next)) {
        let childEnd = index + 1;
        while (childEnd < block.items.length && isIndentedChildItem(current, block.items[childEnd])) {
          childEnd += 1;
        }
        const childLines = block.items
          .slice(index + 1, childEnd)
          .map((item) => formatNarrativeBullet(item.text));
        parts.push(formatNarrativeHeading(current.text));
        if (childLines.length > 0) {
          parts.push(childLines.join("\n"));
        }
        index = childEnd;
        continue;
      }
      parts.push(normalizeNarrativeText(current.text));
      index += 1;
    }
    return parts.join("\n\n");
  }

  function isSectionHeadingNarrativeBlock(block: NarrativeBlock | null | undefined): boolean {
    if (!block || !block.items || block.items.length < 2) {
      return false;
    }
    return isIndentedChildItem(block.items[0], block.items[1]);
  }

  const narrativeStructureApi = {
    renderNarrativeBlock,
    isSectionHeadingNarrativeBlock,
    isCalendarLikeNarrativeBlock
  };

  moduleRegistry.registerModule("narrativeStructure", narrativeStructureApi);
})();
