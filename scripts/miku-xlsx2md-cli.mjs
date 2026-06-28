import fs from "node:fs/promises";

import { runXlsx2mdCli } from "./lib/xlsx2md-cli-runner.mjs";
import { loadXlsx2mdNodeApi } from "./lib/xlsx2md-node-runtime.mjs";

const PACKAGE_JSON_URL = new URL("../package.json", import.meta.url);

async function readPackageVersion() {
  const packageJson = JSON.parse(await fs.readFile(PACKAGE_JSON_URL, "utf8"));
  return packageJson.version || "0.0.0";
}

runXlsx2mdCli({
  argv: process.argv.slice(2),
  loadApi: loadXlsx2mdNodeApi,
  readPackageVersion
}).then((exitCode) => {
  process.exit(exitCode);
}).catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
