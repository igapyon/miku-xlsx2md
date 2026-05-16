import fs from "node:fs/promises";
import path from "node:path";

import {
  RUNTIME_BUNDLE_BASENAME,
  RUNTIME_BUNDLE_JSON,
  RUNTIME_BUNDLE_MJS,
  createRuntimeBundleMetadata,
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
  const { runtimePath } = resolveRuntimeBundlePaths(ROOT);
  const releaseRuntimeName = `${RUNTIME_BUNDLE_BASENAME}-${releaseVersion}.mjs`;
  const releaseMetadataName = `${RUNTIME_BUNDLE_BASENAME}-${releaseVersion}.json`;

  await fs.mkdir(RELEASE_ASSETS_DIR, { recursive: true });
  await copyAsset(runtimePath, releaseRuntimeName);

  const releaseMetadataPath = path.resolve(RELEASE_ASSETS_DIR, releaseMetadataName);
  const releaseMetadata = {
    ...createRuntimeBundleMetadata(packageVersion),
    releaseVersion,
    runtimeAsset: releaseRuntimeName
  };
  await fs.writeFile(releaseMetadataPath, JSON.stringify(releaseMetadata, null, 2) + "\n", "utf8");
  console.log(`[stage:runtime-release] staged ${path.relative(ROOT, releaseMetadataPath)}`);

  console.log(`[stage:runtime-release] source files: ${RUNTIME_BUNDLE_MJS}, ${RUNTIME_BUNDLE_JSON}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
