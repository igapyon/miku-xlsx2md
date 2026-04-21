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

  function renderNarrativeBlock(block: NarrativeBlock): string {
    if (!block.items || block.items.length === 0) {
      return block.lines.map((line) => normalizeNarrativeText(line)).join("\n");
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
    isSectionHeadingNarrativeBlock
  };

  moduleRegistry.registerModule("narrativeStructure", narrativeStructureApi);
})();
