// @vitest-environment jsdom

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { bootRegisteredModule } from "./helpers/xlsx2md-js-loader.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function bootTextEncoding() {
  return bootRegisteredModule(__dirname, ["src/js/text-encoding.js"], "textEncoding");
}

describe("xlsx2md text encoding", () => {
  it("encodes UTF-8 without BOM by default", () => {
    const api = bootTextEncoding();
    expect(Array.from(api.encodeText("A"))).toEqual([0x41]);
  });

  it("encodes UTF-16LE with BOM when requested", () => {
    const api = bootTextEncoding();
    expect(Array.from(api.encodeText("A", { encoding: "utf-16le", bom: "on" }))).toEqual([0xff, 0xfe, 0x41, 0x00]);
  });

  it("encodes UTF-32BE without BOM", () => {
    const api = bootTextEncoding();
    expect(Array.from(api.encodeText("A", { encoding: "utf-32be", bom: "off" }))).toEqual([0x00, 0x00, 0x00, 0x41]);
  });

  it("encodes shift_jis in the Node-backed runtime", () => {
    const api = bootTextEncoding();
    expect(Array.from(api.encodeText("あA", { encoding: "shift_jis", bom: "off" }))).toEqual([0x82, 0xa0, 0x41]);
  });

  it("rejects BOM for shift_jis", () => {
    const api = bootTextEncoding();
    expect(() => api.encodeText("A", { encoding: "shift_jis", bom: "on" })).toThrow("BOM cannot be enabled for shift_jis.");
  });
});
