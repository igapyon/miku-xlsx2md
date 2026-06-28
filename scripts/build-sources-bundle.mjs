import fs from "node:fs/promises";
import path from "node:path";
import zlib from "node:zlib";

import { resolveRuntimeBundlePaths } from "./lib/xlsx2md-runtime-bundle.mjs";

const ROOT = process.cwd();
const INCLUDE_PREFIXES = [
  ".github/workflows/",
  "docs/",
  "scripts/",
  "src/",
  "tests/"
];
const INCLUDE_FILES = [
  "CONTRIBUTING.md",
  "LICENSE",
  "NOTICE",
  "README.md",
  "THIRD_PARTY_NOTICES.md",
  "TODO.md",
  "package-lock.json",
  "package.json"
];

function shouldInclude(relPath) {
  return INCLUDE_FILES.includes(relPath) || INCLUDE_PREFIXES.some((prefix) => relPath.startsWith(prefix));
}

async function listSourceFiles(dir, prefix = "") {
  const entries = await fs.readdir(path.resolve(dir, prefix), { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const relPath = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.name === ".DS_Store" || entry.name === "node_modules" || entry.name === "dist" ||
        entry.name === "bundle" || entry.name === "release-assets" || entry.name === "workplace") {
      continue;
    }
    if (entry.isDirectory()) {
      if (shouldInclude(`${relPath}/`) || INCLUDE_PREFIXES.some((prefixValue) => prefixValue.startsWith(`${relPath}/`))) {
        files.push(...await listSourceFiles(dir, relPath));
      }
      continue;
    }
    if (entry.isFile() && shouldInclude(relPath)) {
      files.push(relPath);
    }
  }
  return files.sort();
}

function writeString(buffer, offset, value, length) {
  buffer.write(value.slice(0, length), offset, length, "utf8");
}

function writeOctal(buffer, offset, length, value) {
  const text = value.toString(8).padStart(length - 1, "0") + "\0";
  buffer.write(text, offset, length, "ascii");
}

function createTarHeader(name, size, mode) {
  const header = Buffer.alloc(512, 0);
  writeString(header, 0, name, 100);
  writeOctal(header, 100, 8, mode);
  writeOctal(header, 108, 8, 0);
  writeOctal(header, 116, 8, 0);
  writeOctal(header, 124, 12, size);
  writeOctal(header, 136, 12, 0);
  header.fill(" ", 148, 156);
  header[156] = "0".charCodeAt(0);
  writeString(header, 257, "ustar", 6);
  writeString(header, 263, "00", 2);
  let checksum = 0;
  for (const byte of header) {
    checksum += byte;
  }
  const checksumText = checksum.toString(8).padStart(6, "0") + "\0 ";
  header.write(checksumText, 148, 8, "ascii");
  return header;
}

function padToBlock(data) {
  const remainder = data.length % 512;
  return remainder === 0 ? Buffer.alloc(0) : Buffer.alloc(512 - remainder, 0);
}

async function createTarGz(files) {
  const parts = [];
  for (const relPath of files) {
    const data = await fs.readFile(path.resolve(ROOT, relPath));
    const archiveName = `miku-xlsx2md/${relPath}`;
    if (Buffer.byteLength(archiveName) > 100) {
      throw new Error(`Source archive path is too long for ustar header: ${archiveName}`);
    }
    parts.push(createTarHeader(archiveName, data.length, 0o644), data, padToBlock(data));
  }
  parts.push(Buffer.alloc(1024, 0));
  return zlib.gzipSync(Buffer.concat(parts), { level: 9, mtime: 0 });
}

async function main() {
  const { bundleDir, sourceArchivePath } = resolveRuntimeBundlePaths(ROOT);
  const files = await listSourceFiles(ROOT);
  if (files.length === 0) {
    throw new Error("No source files selected for source bundle.");
  }
  await fs.mkdir(bundleDir, { recursive: true });
  await fs.writeFile(sourceArchivePath, await createTarGz(files));
  console.log(`[build:bundle] generated ${path.relative(ROOT, sourceArchivePath)} (${files.length} files)`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
