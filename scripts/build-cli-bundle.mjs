import fs from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";

import {
  getPackageVersion,
  readPackageJson,
  resolveRuntimeBundlePaths
} from "./lib/xlsx2md-runtime-bundle.mjs";
import { XLSX2MD_CORE_JS_ORDER } from "./lib/xlsx2md-module-order.mjs";

const ROOT = process.cwd();

async function createCliBundleEntry(version) {
  const coreSources = [];
  for (const relPath of XLSX2MD_CORE_JS_ORDER) {
    coreSources.push({
      path: relPath,
      source: await fs.readFile(path.resolve(ROOT, relPath), "utf8")
    });
  }

  return `import { Blob as NodeBlob } from "node:buffer";
import { createRequire } from "node:module";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";
import {
  DOMParser as XmldomParser,
  XMLSerializer as XmldomSerializer,
  Node as XmldomNode,
  Document as XmldomDocument,
  Element as XmldomElement
} from "@xmldom/xmldom";
import { runXlsx2mdCli } from "../scripts/lib/xlsx2md-cli-runner.mjs";

const XLSX2MD_RUNTIME_CORE_SOURCES = ${JSON.stringify(coreSources)};
const PACKAGE_VERSION = ${JSON.stringify(version)};
const nodeRequire = createRequire(import.meta.url);
let cachedApi = null;

function installNodeDomGlobals() {
  if (typeof globalThis.DOMParser !== "function") {
    globalThis.DOMParser = XmldomParser;
    globalThis.Node = XmldomNode;
    globalThis.Document = XmldomDocument;
    globalThis.Element = XmldomElement;
    globalThis.XMLSerializer ??= XmldomSerializer;
  }
  if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
    globalThis.Blob = NodeBlob;
  }
  globalThis.DecompressionStream ??= NodeDecompressionStream;
  globalThis.__xlsx2mdNodeRequire ??= nodeRequire;
}

function loadBundledXlsx2mdApi() {
  if (cachedApi) {
    return cachedApi;
  }
  installNodeDomGlobals();
  delete globalThis.__xlsx2mdModuleRegistry;
  delete globalThis.__xlsx2mdModuleRegistryStore;
  delete globalThis.getXlsx2mdModuleRegistry;
  for (const entry of XLSX2MD_RUNTIME_CORE_SOURCES) {
    new Function(entry.source)();
  }
  const api = globalThis.__xlsx2mdModuleRegistry?.getModule("xlsx2md");
  if (!api) {
    throw new Error("xlsx2md bundled CLI API failed to initialize.");
  }
  cachedApi = api;
  return api;
}

runXlsx2mdCli({
  argv: process.argv.slice(2),
  loadApi: loadBundledXlsx2mdApi,
  readPackageVersion: async () => PACKAGE_VERSION
}).then((exitCode) => {
  process.exit(exitCode);
}).catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
`;
}

async function main() {
  const packageJson = await readPackageJson(ROOT);
  const version = getPackageVersion(packageJson);
  const { bundleDir, cliPath } = resolveRuntimeBundlePaths(ROOT);
  const entryPath = path.resolve(bundleDir, ".miku-xlsx2md-cli-entry.mjs");

  await fs.mkdir(bundleDir, { recursive: true });
  await fs.writeFile(entryPath, await createCliBundleEntry(version), "utf8");
  await build({
    entryPoints: [entryPath],
    outfile: cliPath,
    bundle: true,
    platform: "node",
    format: "esm",
    target: "node20",
    banner: {
      js: "#!/usr/bin/env node"
    },
    legalComments: "none"
  });
  await fs.rm(entryPath, { force: true });
  await fs.chmod(cliPath, 0o755);
  console.log(`[build:bundle] generated ${path.relative(ROOT, cliPath)}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
