import fs from "node:fs/promises";
import path from "node:path";

import { XLSX2MD_CORE_JS_ORDER } from "./xlsx2md-module-order.mjs";

export const PRODUCT_BUNDLE_BASENAME = "miku-xlsx2md";
export const RUNTIME_BUNDLE_BASENAME = "miku-xlsx2md-runtime";
export const RUNTIME_BUNDLE_DIR = "bundle";
export const CLI_BUNDLE_MJS = `${PRODUCT_BUNDLE_BASENAME}.mjs`;
export const RUNTIME_BUNDLE_MJS = `${RUNTIME_BUNDLE_BASENAME}.mjs`;
export const RUNTIME_BUNDLE_JSON = `${RUNTIME_BUNDLE_BASENAME}.json`;
export const SOURCE_BUNDLE_TGZ = `${PRODUCT_BUNDLE_BASENAME}-sources.tgz`;

export async function readPackageJson(rootDir) {
  return JSON.parse(await fs.readFile(path.resolve(rootDir, "package.json"), "utf8"));
}

export function getPackageVersion(packageJson) {
  const version = packageJson.version;
  if (!version) {
    throw new Error("package.json version is required to build the runtime bundle.");
  }
  return version;
}

export async function createRuntimeBundleSource(rootDir, version) {
  const coreSources = [];
  for (const relPath of XLSX2MD_CORE_JS_ORDER) {
    coreSources.push({
      path: relPath,
      source: await fs.readFile(path.resolve(rootDir, relPath), "utf8")
    });
  }

  return `/*
 * miku-xlsx2md runtime bundle
 * Source repository: https://github.com/igapyon/miku-xlsx2md
 * Version: ${version}
 */
const XLSX2MD_RUNTIME_CORE_SOURCES = ${JSON.stringify(coreSources)};
let cachedApi = null;

export const version = ${JSON.stringify(version)};
export const embeddedCorePaths = XLSX2MD_RUNTIME_CORE_SOURCES.map((entry) => entry.path);

export function loadXlsx2mdRuntime(options = {}) {
  if (cachedApi && !options.reset) {
    return cachedApi;
  }

  if (options.reset) {
    delete globalThis.__xlsx2mdModuleRegistry;
    delete globalThis.__xlsx2mdModuleRegistryStore;
    delete globalThis.getXlsx2mdModuleRegistry;
  }

  for (const entry of XLSX2MD_RUNTIME_CORE_SOURCES) {
    new Function(entry.source)();
  }

  const api = globalThis.__xlsx2mdModuleRegistry?.getModule("xlsx2md");
  if (!api) {
    throw new Error("xlsx2md runtime API failed to initialize.");
  }

  cachedApi = api;
  return api;
}

export default loadXlsx2mdRuntime;
`;
}

export function createRuntimeBundleMetadata(version) {
  return {
    repository: "https://github.com/igapyon/miku-xlsx2md",
    runtimeVersion: version,
    runtimeAsset: RUNTIME_BUNDLE_MJS,
    embeddedCorePaths: XLSX2MD_CORE_JS_ORDER
  };
}

export function resolveRuntimeBundlePaths(rootDir) {
  const bundleDir = path.resolve(rootDir, RUNTIME_BUNDLE_DIR);
  return {
    bundleDir,
    cliPath: path.resolve(bundleDir, CLI_BUNDLE_MJS),
    runtimePath: path.resolve(bundleDir, RUNTIME_BUNDLE_MJS),
    metadataPath: path.resolve(bundleDir, RUNTIME_BUNDLE_JSON),
    sourceArchivePath: path.resolve(bundleDir, SOURCE_BUNDLE_TGZ)
  };
}
