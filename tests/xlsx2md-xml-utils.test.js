// @vitest-environment jsdom

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootRegisteredModule } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootXmlUtils() {
  return bootRegisteredModule(__dirname, [
    "src/js/runtime-env.js",
    "src/js/xml-utils.js"
  ], "xmlUtils");
}

describe("xlsx2md xml utils", () => {
  it("finds elements by local name across namespaces", () => {
    const api = bootXmlUtils();
    const doc = api.xmlToDocument('<root xmlns:a="urn:test"><a:item>v1</a:item><a:item>v2</a:item></root>');

    expect(api.getElementsByLocalName(doc, "item").map((node) => node.textContent)).toEqual(["v1", "v2"]);
    expect(api.getFirstChildByLocalName(doc, "item")?.textContent).toBe("v1");
  });

  it("finds only direct children for a local name", () => {
    const api = bootXmlUtils();
    const doc = api.xmlToDocument("<root><item><item>nested</item></item><item>direct</item></root>");
    const root = doc.documentElement;

    expect(api.getDirectChildByLocalName(root, "item")?.textContent).toBe("nested");
  });

  it("decodes utf-8 bytes and normalizes CRLF text content", () => {
    const api = bootXmlUtils();
    const encoded = new TextEncoder().encode("<root>line1\r\nline2</root>");
    const doc = api.xmlToDocument(api.decodeXmlText(encoded));

    expect(api.getTextContent(doc.documentElement)).toBe("line1\nline2");
  });
});
