// @vitest-environment node

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");

describe("xlsx2md runtime release staging", () => {
  it("stages product and runtime mjs assets without release json metadata", () => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), "xlsx2md-release-"));
    fs.mkdirSync(path.join(tempRoot, "bundle"), { recursive: true });
    fs.mkdirSync(path.join(tempRoot, "release-assets"), { recursive: true });
    fs.writeFileSync(
      path.join(tempRoot, "package.json"),
      JSON.stringify({ name: "miku-xlsx2md", version: "9.8.7" }, null, 2) + "\n"
    );
    fs.writeFileSync(path.join(tempRoot, "bundle", "miku-xlsx2md.mjs"), "console.log('ok');\n");
    fs.writeFileSync(path.join(tempRoot, "bundle", "miku-xlsx2md-runtime.mjs"), "export const ok = true;\n");
    fs.writeFileSync(path.join(tempRoot, "bundle", "miku-xlsx2md-sources.tgz"), "sources\n");
    fs.writeFileSync(path.join(tempRoot, "release-assets", "stale.json"), "{}\n");

    const result = spawnSync("node", [path.join(ROOT, "scripts", "stage-runtime-release-assets.mjs")], {
      cwd: tempRoot,
      env: {
        ...process.env,
        TAG_NAME: "v9.8.7"
      },
      encoding: "utf8"
    });

    expect(result.status).toBe(0);
    expect(fs.readdirSync(path.join(tempRoot, "release-assets")).sort()).toEqual([
      "miku-xlsx2md-9.8.7.mjs",
      "miku-xlsx2md-runtime-9.8.7.mjs",
      "miku-xlsx2md-sources-9.8.7.tgz"
    ]);
  });
});
