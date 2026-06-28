import fs from "node:fs/promises";
import path from "node:path";

import {
  PRODUCT_BUNDLE_BASENAME,
  SOURCE_BUNDLE_TGZ,
  RUNTIME_BUNDLE_JSON,
  RUNTIME_BUNDLE_BASENAME,
  RUNTIME_BUNDLE_MJS,
  getPackageVersion,
  readPackageJson,
  resolveRuntimeBundlePaths
} from "./lib/xlsx2md-runtime-bundle.mjs";

const ROOT = process.cwd();
const RELEASE_ASSETS_DIR = path.resolve(ROOT, "release-assets");

function resolveReleaseVersion(packageVersion) {
  const tagName = process.env.TAG_NAME || "";
  if (!tagName) {
    return packageVersion;
  }
  if (!tagName.startsWith("v")) {
    throw new Error(`Release tag must start with "v": ${tagName}`);
  }
  const version = tagName.slice(1);
  if (version !== packageVersion && !version.startsWith(`${packageVersion}.`)) {
    throw new Error(
      `Release tag version (${version}) must match package.json version (${packageVersion}) or add a dot suffix.`
    );
  }
  return version;
}

async function copyAsset(sourcePath, targetName) {
  const targetPath = path.resolve(RELEASE_ASSETS_DIR, targetName);
  await fs.copyFile(sourcePath, targetPath);
  console.log(`[stage:runtime-release] staged ${path.relative(ROOT, targetPath)}`);
}

async function main() {
  const packageJson = await readPackageJson(ROOT);
  const packageVersion = getPackageVersion(packageJson);
  const releaseVersion = resolveReleaseVersion(packageVersion);
  const { cliPath, runtimePath, sourceArchivePath } = resolveRuntimeBundlePaths(ROOT);
  const releaseCliName = `${PRODUCT_BUNDLE_BASENAME}-${releaseVersion}.mjs`;
  const releaseRuntimeName = `${RUNTIME_BUNDLE_BASENAME}-${releaseVersion}.mjs`;
  const releaseSourcesName = `${PRODUCT_BUNDLE_BASENAME}-sources-${releaseVersion}.tgz`;

  await fs.rm(RELEASE_ASSETS_DIR, { recursive: true, force: true });
  await fs.mkdir(RELEASE_ASSETS_DIR, { recursive: true });
  await copyAsset(cliPath, releaseCliName);
  await copyAsset(runtimePath, releaseRuntimeName);
  await copyAsset(sourceArchivePath, releaseSourcesName);

  console.log(`[stage:runtime-release] source files: ${PRODUCT_BUNDLE_BASENAME}.mjs, ${RUNTIME_BUNDLE_MJS}, ${SOURCE_BUNDLE_TGZ}, ${RUNTIME_BUNDLE_JSON}`);
  console.log("[stage:runtime-release] metadata json is a local build artifact and is not staged for GitHub Release");
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
