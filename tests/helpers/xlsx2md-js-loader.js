import { Blob as NodeBlob } from "node:buffer";
import { readFileSync } from "node:fs";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";

import { XLSX2MD_CORE_JS_ORDER } from "../../scripts/lib/xlsx2md-module-order.mjs";
import { loadModuleRegistry } from "./module-registry.js";

const FORMULA_RUNTIME_MODULES = new Set([
  "src/js/formula/tokenizer.js",
  "src/js/formula/parser.js",
  "src/js/formula/evaluator.js"
]);

export function installWorkbookRuntimeGlobals() {
  if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
    globalThis.Blob = NodeBlob;
  }
  globalThis.DecompressionStream ??= NodeDecompressionStream;
}

export function loadJsModule(testDir, relPath) {
  const code = readFileSync(path.resolve(testDir, "..", relPath), "utf8");
  new Function(code)();
}

export function loadJsModules(testDir, relPaths) {
  relPaths.forEach((relPath) => loadJsModule(testDir, relPath));
}

export function resetModuleRegistryStore() {
  delete globalThis.__xlsx2mdModuleRegistry;
  delete globalThis.__xlsx2mdModuleRegistryStore;
}

export function bootRegisteredModule(testDir, relPaths, moduleName) {
  document.body.innerHTML = "";
  installWorkbookRuntimeGlobals();
  resetModuleRegistryStore();
  loadModuleRegistry(testDir);
  loadJsModules(testDir, relPaths);
  return globalThis.__xlsx2mdModuleRegistry.getModule(moduleName);
}

export function bootXlsx2mdCore(testDir, options = {}) {
  document.body.innerHTML = "";
  installWorkbookRuntimeGlobals();
  resetModuleRegistryStore();
  loadModuleRegistry(testDir);
  loadJsModules(
    testDir,
    XLSX2MD_CORE_JS_ORDER.filter((relPath) => (
      relPath !== "src/js/module-registry.js" && relPath !== "src/js/module-registry-access.js"
      && (options.includeFormulaRuntime === true || !FORMULA_RUNTIME_MODULES.has(relPath))
    ))
  );
  return globalThis.__xlsx2mdModuleRegistry.getModule("xlsx2md");
}

export function bootRichTextParser(testDir) {
  document.body.innerHTML = "";
  resetModuleRegistryStore();
  loadModuleRegistry(testDir);
  loadJsModules(testDir, [
    "src/js/markdown-normalize.js",
    "src/js/markdown-escape.js",
    "src/js/rich-text-parser.js"
  ]);
  return globalThis.__xlsx2mdModuleRegistry.getModule("richTextParser")
    .createRichTextParserApi({
      normalizeMarkdownText: (text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\t/g, " ")
    });
}

export function bootRichTextRenderer(testDir) {
  document.body.innerHTML = "";
  resetModuleRegistryStore();
  loadModuleRegistry(testDir);
  loadJsModules(testDir, [
    "src/js/markdown-normalize.js",
    "src/js/markdown-escape.js",
    "src/js/rich-text-parser.js",
    "src/js/rich-text-plain-formatter.js",
    "src/js/rich-text-github-formatter.js",
    "src/js/rich-text-renderer.js"
  ]);
  return globalThis.__xlsx2mdModuleRegistry.getModule("richTextRenderer")
    .createRichTextRendererApi({
      normalizeMarkdownText: (text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\t/g, " ")
    });
}

export function bootSheetMarkdownModule(testDir) {
  document.body.innerHTML = "";
  resetModuleRegistryStore();
  loadModuleRegistry(testDir);
  loadJsModules(testDir, [
    "src/js/markdown-normalize.js",
    "src/js/markdown-escape.js",
    "src/js/markdown-table-escape.js",
    "src/js/markdown-options.js",
    "src/js/rich-text-parser.js",
    "src/js/rich-text-plain-formatter.js",
    "src/js/rich-text-github-formatter.js",
    "src/js/rich-text-renderer.js",
    "src/js/sheet-markdown.js"
  ]);
  return globalThis.__xlsx2mdModuleRegistry.getModule("sheetMarkdown");
}
