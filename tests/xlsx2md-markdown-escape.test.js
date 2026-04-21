// @vitest-environment jsdom

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootRegisteredModule } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootMarkdownEscape() {
  return bootRegisteredModule(__dirname, [
    "src/js/markdown-normalize.js",
    "src/js/markdown-escape.js"
  ], "markdownEscape");
}

describe("xlsx2md markdown escape", () => {
  it("escapes inline markdown control characters and html-sensitive brackets", () => {
    const api = bootMarkdownEscape();

    expect(api.escapeMarkdownLiteralText("a*b _c_ [x](y) ![z](w) <tag> a | b `q` ~"))
      .toBe("a\\*b \\_c\\_ \\[x\\]\\(y\\) \\!\\[z\\]\\(w\\) &lt;tag&gt; a \\| b \\`q\\` \\~");
    expect(api.escapeMarkdownLiteralParts("a*b <tag>")).toEqual([
      { kind: "text", text: "a", rawText: "a" },
      { kind: "escaped", text: "\\*", rawText: "*" },
      { kind: "text", text: "b ", rawText: "b " },
      { kind: "escaped", text: "&lt;", rawText: "<" },
      { kind: "text", text: "tag", rawText: "tag" },
      { kind: "escaped", text: "&gt;", rawText: ">" }
    ]);
  });

  it("escapes line-start markdown markers line by line", () => {
    const api = bootMarkdownEscape();

    expect(api.escapeMarkdownLineStart("# h\n- item\r\n1. num\r> quote"))
      .toBe("\\# h\n\\- item\n1\\. num\n\\> quote");
    expect(api.escapeMarkdownLiteralText("# h\n- item\n1. num\n> quote"))
      .toBe("\\# h\n\\- item\n1\\. num\n&gt; quote");
  });

  it("escapes additional list markers and ampersands", () => {
    const api = bootMarkdownEscape();

    expect(api.escapeMarkdownLiteralText("+ plus\n* star\na & b"))
      .toBe("\\+ plus\n\\* star\na &amp; b");
    expect(api.escapeMarkdownLiteralParts("a & b")).toEqual([
      { kind: "text", text: "a ", rawText: "a " },
      { kind: "escaped", text: "&amp;", rawText: "&" },
      { kind: "text", text: " b", rawText: " b" }
    ]);
  });
});
