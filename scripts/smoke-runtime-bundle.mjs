import { pathToFileURL } from "node:url";

import {
  getPackageVersion,
  readPackageJson,
  resolveRuntimeBundlePaths
} from "./lib/xlsx2md-runtime-bundle.mjs";

const ROOT = process.cwd();

async function main() {
  const packageJson = await readPackageJson(ROOT);
  const expectedVersion = getPackageVersion(packageJson);
  const { runtimePath } = resolveRuntimeBundlePaths(ROOT);
  const runtime = await import(`${pathToFileURL(runtimePath).href}?smoke=${Date.now()}`);

  if (runtime.version !== expectedVersion) {
    throw new Error(`Runtime version mismatch: expected ${expectedVersion}, got ${runtime.version}`);
  }
  if (!Array.isArray(runtime.embeddedCorePaths) || runtime.embeddedCorePaths.length === 0) {
    throw new Error("Runtime bundle does not expose embeddedCorePaths.");
  }

  const api = runtime.loadXlsx2mdRuntime({ reset: true });
  if (!api || typeof api.parseWorkbook !== "function") {
    throw new Error("Runtime bundle did not expose xlsx2md.parseWorkbook.");
  }
  if (!api.markdownOptions?.OUTPUT_MODES?.includes("display")) {
    throw new Error("Runtime bundle did not expose markdown option definitions.");
  }

  console.log(`[smoke:runtime] ${runtime.version} ok`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
