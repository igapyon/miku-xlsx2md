import fs from "node:fs";
import path from "node:path";
import { transform } from "esbuild";
import { XLSX2MD_CORE_TS_ORDER } from "./lib/xlsx2md-module-order.mjs";

const ROOT = process.cwd();
const CORE_VENDOR_PATH = "src/vendor/miku-ms-office-core-0.5.0.1.mjs";

const tsModule = await loadTypeScriptModule();
await generateMsOfficeCoreAdapter();
transpileTypeScriptFiles(XLSX2MD_CORE_TS_ORDER, tsModule);
console.log("[build:miku-xlsx2md] generated core JavaScript files under dist/js");

async function loadTypeScriptModule() {
  try {
    const module = await import("typescript");
    return module.default || module;
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new Error(
      "TypeScript is required for build. Install dependencies before running `npm run build`.\n" +
      `Cause: ${reason}`
    );
  }
}

async function generateMsOfficeCoreAdapter() {
  const vendorPath = path.resolve(ROOT, CORE_VENDOR_PATH);
  const jsPath = path.resolve(ROOT, "dist/js/ms-office-core.js");
  const source = fs.readFileSync(vendorPath, "utf8");
  const result = await transform(source, {
    format: "iife",
    globalName: "__mikuMsOfficeCoreRelease",
    target: "es2019"
  });
  const adapter = `${result.code}
(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  moduleRegistry.registerModule("msOfficeCore", __mikuMsOfficeCoreRelease);
})();
`;

  fs.mkdirSync(path.dirname(jsPath), { recursive: true });
  fs.writeFileSync(jsPath, adapter, "utf8");
}

function transpileTypeScriptFiles(tsOrder, tsModule) {
  for (const relTsPath of tsOrder) {
    const tsPath = path.resolve(ROOT, relTsPath);
    const jsPath = path.resolve(
      ROOT,
      relTsPath.replace("src/ts/", "dist/js/").replace(/\.ts$/, ".js")
    );

    const source = fs.readFileSync(tsPath, "utf8");
    const result = tsModule.transpileModule(source, {
      compilerOptions: {
        target: tsModule.ScriptTarget.ES2019,
        module: tsModule.ModuleKind.None,
        lib: ["ES2020", "DOM"],
        strict: false,
        skipLibCheck: true
      },
      reportDiagnostics: true,
      fileName: tsPath
    });

    if (result.diagnostics && result.diagnostics.length > 0) {
      const errors = result.diagnostics
        .filter((diagnostic) => diagnostic.category === tsModule.DiagnosticCategory.Error)
        .map((diagnostic) => tsModule.flattenDiagnosticMessageText(diagnostic.messageText, "\n"));
      if (errors.length > 0) {
        throw new Error(`TypeScript transpile error in ${relTsPath}:\n${errors.join("\n")}`);
      }
    }
    const outputText = result.outputText;
    fs.mkdirSync(path.dirname(jsPath), { recursive: true });
    fs.writeFileSync(jsPath, outputText, "utf8");
  }
}
