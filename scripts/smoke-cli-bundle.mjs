import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

import {
  getPackageVersion,
  readPackageJson,
  resolveRuntimeBundlePaths
} from "./lib/xlsx2md-runtime-bundle.mjs";

const ROOT = process.cwd();

function runNode(args) {
  return spawnSync(process.execPath, args, {
    cwd: ROOT,
    encoding: "utf8"
  });
}

function assertSuccess(result, label) {
  if (result.status !== 0) {
    throw new Error(`${label} failed with exit ${result.status}\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
  }
}

async function main() {
  const packageJson = await readPackageJson(ROOT);
  const expectedVersion = getPackageVersion(packageJson);
  const { cliPath } = resolveRuntimeBundlePaths(ROOT);

  const versionResult = runNode([cliPath, "--version"]);
  assertSuccess(versionResult, "CLI bundle --version");
  if (versionResult.stdout.trim() !== expectedVersion) {
    throw new Error(`CLI bundle version mismatch: expected ${expectedVersion}, got ${versionResult.stdout.trim()}`);
  }

  const helpResult = runNode([cliPath, "--help"]);
  assertSuccess(helpResult, "CLI bundle --help");
  if (!helpResult.stdout.includes("Usage:") || !helpResult.stdout.includes("--version")) {
    throw new Error("CLI bundle help output does not contain expected usage text.");
  }

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "xlsx2md-cli-bundle-"));
  const outPath = path.join(tempDir, "sample.md");
  const convertResult = runNode([
    cliPath,
    path.resolve(ROOT, "tests/fixtures/xlsx2md-basic-sample01.xlsx"),
    "--out",
    outPath
  ]);
  assertSuccess(convertResult, "CLI bundle fixture conversion");
  const markdown = await fs.readFile(outPath, "utf8");
  if (!markdown.includes("# Book: xlsx2md-basic-sample01.xlsx")) {
    throw new Error("CLI bundle conversion did not produce expected Markdown.");
  }

  console.log(`[smoke:bundle] ${expectedVersion} ok`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
