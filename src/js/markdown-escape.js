/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();
    function escapeMarkdownLineStartSegment(text) {
        return String(text || "")
            .replace(/^(\s*)([#>])/u, "$1\\$2")
            .replace(/^(\s*)([-+*])(\s+)/u, "$1\\$2$3")
            .replace(/^(\s*)(\d+)\.(\s+)/u, "$1$2\\.$3");
    }
    function escapeMarkdownLineStart(text) {
        return markdownNormalizeHelper
            .normalizeMarkdownNewlines(text)
            .split("\n")
            .map((line) => escapeMarkdownLineStartSegment(line))
            .join("\n");
    }
    function getEscapedMarkdownLiteralText(ch, atLineStart, nextChar) {
        if (ch === "\\")
            return "\\\\";
        if (ch === "&")
            return "&amp;";
        if (ch === "<")
            return "&lt;";
        if (ch === ">")
            return "&gt;";
        if (/[`*_{}\[\]()!|~]/.test(ch))
            return `\\${ch}`;
        if (atLineStart && ch === "#")
            return `\\${ch}`;
        if (atLineStart && /[-+*]/.test(ch) && /\s/u.test(nextChar))
            return `\\${ch}`;
        return null;
    }
    function parseOrderedListMarker(source, index, atLineStart) {
        if (!atLineStart || !/\d/u.test(source[index] || "")) {
            return null;
        }
        let digits = source[index];
        let cursor = index + 1;
        while (cursor < source.length && /\d/u.test(source[cursor])) {
            digits += source[cursor];
            cursor += 1;
        }
        if (source[cursor] !== "." || !/\s/u.test(source[cursor + 1] || "")) {
            return null;
        }
        return {
            digits,
            dotIndex: cursor
        };
    }
    function escapeMarkdownLiteralParts(text) {
        const source = String(text || "");
        const parts = [];
        let buffer = "";
        function pushTextBuffer() {
            if (!buffer)
                return;
            parts.push({ kind: "text", text: buffer, rawText: buffer });
            buffer = "";
        }
        function pushEscaped(textValue, rawText) {
            pushTextBuffer();
            if (!textValue)
                return;
            parts.push({ kind: "escaped", text: textValue, rawText });
        }
        for (let index = 0; index < source.length; index += 1) {
            const ch = source[index];
            const atLineStart = index === 0;
            const next = source[index + 1] || "";
            const escapedText = getEscapedMarkdownLiteralText(ch, atLineStart, next);
            if (escapedText) {
                pushEscaped(escapedText, ch);
                continue;
            }
            const orderedListMarker = parseOrderedListMarker(source, index, atLineStart);
            if (orderedListMarker) {
                pushTextBuffer();
                parts.push({ kind: "text", text: orderedListMarker.digits, rawText: orderedListMarker.digits });
                parts.push({ kind: "escaped", text: "\\.", rawText: "." });
                index = orderedListMarker.dotIndex;
                continue;
            }
            buffer += ch;
        }
        pushTextBuffer();
        return parts;
    }
    function escapeMarkdownLiteralText(text) {
        return markdownNormalizeHelper
            .normalizeMarkdownNewlines(text)
            .split("\n")
            .map((line) => escapeMarkdownLiteralParts(line).map((part) => part.text).join(""))
            .join("\n");
    }
    const markdownEscapeApi = {
        escapeMarkdownLineStart,
        escapeMarkdownLiteralParts,
        escapeMarkdownLiteralText
    };
    moduleRegistry.registerModule("markdownEscape", markdownEscapeApi);
})();
