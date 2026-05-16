import fs from "node:fs/promises";
import path from "node:path";

import {
  createRuntimeBundleMetadata,
  createRuntimeBundleSource,
  getPackageVersion,
  readPackageJson,
  resolveRuntimeBundlePaths
} from "./lib/xlsx2md-runtime-bundle.mjs";

const ROOT = process.cwd();

async function main() {
  const packageJson = await readPackageJson(ROOT);
  const version = getPackageVersion(packageJson);
  const { bundleDir, runtimePath, metadataPath } = resolveRuntimeBundlePaths(ROOT);

  await fs.mkdir(bundleDir, { recursive: true });
  await fs.writeFile(runtimePath, await createRuntimeBundleSource(ROOT, version), "utf8");
  await fs.writeFile(
    metadataPath,
    JSON.stringify(createRuntimeBundleMetadata(version), null, 2) + "\n",
    "utf8"
  );

  console.log(`[build:runtime] generated ${path.relative(ROOT, runtimePath)}`);
  console.log(`[build:runtime] generated ${path.relative(ROOT, metadataPath)}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
