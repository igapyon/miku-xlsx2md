import fs from "node:fs/promises";
import path from "node:path";

import { loadXlsx2mdNodeApi } from "./lib/xlsx2md-node-runtime.mjs";

const SHAPE_DETAILS_MODES = ["include", "exclude"];
const ENCODINGS = ["utf-8", "shift_jis", "utf-16le", "utf-16be", "utf-32le", "utf-32be"];
const BOM_MODES = ["off", "on"];
const FLAG_OPTIONS = {
  "--help"(options) {
    options.help = true;
  },
  "--include-shape-details"(options) {
    options.includeShapeDetails = true;
  },
  "--no-header-row"(options) {
    options.treatFirstRowAsHeader = false;
  },
  "--no-trim-text"(options) {
    options.trimText = false;
  },
  "--keep-empty-rows"(options) {
    options.removeEmptyRows = false;
  },
  "--keep-empty-columns"(options) {
    options.removeEmptyColumns = false;
  },
  "--summary"(options) {
    options.summary = true;
  }
};

function printHelp() {
  console.log(`Usage:
  node scripts/miku-xlsx2md-cli.mjs <input.xlsx> [options]

Options:
  --out <file>                  Write combined Markdown to this file
  --zip <file>                  Write ZIP export to this file
  --encoding <value>            utf-8 | shift_jis | utf-16le | utf-16be | utf-32le | utf-32be (default: utf-8)
  --bom <value>                 off | on (default: off; shift_jis does not allow on)
  --output-mode <mode>          display | raw | both (default: display)
  --formatting-mode <mode>      plain | github (default: github)
  --table-detection-mode <mode> balanced | border | planner-aware (default: balanced)
  --shape-details <mode>        include | exclude (default: exclude)
  --include-shape-details       Alias for --shape-details include
  --no-header-row               Do not treat the first row as a table header
  --no-trim-text                Preserve surrounding whitespace
  --keep-empty-rows             Keep empty rows
  --keep-empty-columns          Keep empty columns
  --summary                     Print per-sheet summary to stdout
  --help                        Show this help and exit

GUI-aligned defaults:
  output-mode=display, formatting-mode=github, table-detection-mode=balanced, shape-details=exclude

Exit codes:
  0                             Success
  1                             Error
`);
}

function normalizeEnumOption(value, allowedValues, label, aliases = {}) {
  const normalized = aliases[value] || value;
  if (!allowedValues.includes(normalized)) {
    throw new Error(`Invalid ${label}: ${value}`);
  }
  return normalized;
}

function createValueOptions(markdownOptions) {
  return {
    "--out": {
      apply(options, value) {
        options.outPath = value;
      }
    },
    "--zip": {
      apply(options, value) {
        options.zipPath = value;
      }
    },
    "--output-mode": {
      validate: (value) => normalizeEnumOption(value, markdownOptions.OUTPUT_MODES, "output mode"),
      apply(options, value) {
        options.outputMode = value;
      }
    },
    "--formatting-mode": {
      validate: (value) => normalizeEnumOption(value, markdownOptions.FORMATTING_MODES, "formatting mode"),
      apply(options, value) {
        options.formattingMode = value;
      }
    },
    "--table-detection-mode": {
      validate: (value) => normalizeEnumOption(
        value,
        markdownOptions.TABLE_DETECTION_MODES,
        "table detection mode",
        markdownOptions.TABLE_DETECTION_MODE_ALIASES
      ),
      apply(options, value) {
        options.tableDetectionMode = value;
      }
    },
    "--shape-details": {
      validate: (value) => normalizeEnumOption(value, SHAPE_DETAILS_MODES, "shape details mode"),
      apply(options, value) {
        options.includeShapeDetails = value === "include";
      }
    },
    "--encoding": {
      validate: (value) => normalizeEnumOption(value, ENCODINGS, "encoding"),
      apply(options, value) {
        options.encoding = value;
      }
    },
    "--bom": {
      validate: (value) => normalizeEnumOption(value, BOM_MODES, "BOM mode"),
      apply(options, value) {
        options.bom = value;
      }
    }
  };
}

function parseArgs(argv, markdownOptions) {
  const valueOptions = createValueOptions(markdownOptions);
  const options = {
    treatFirstRowAsHeader: true,
    trimText: true,
    removeEmptyRows: true,
    removeEmptyColumns: true,
    includeShapeDetails: false,
    outputMode: "display",
    formattingMode: "github",
    tableDetectionMode: "balanced",
    encoding: "utf-8",
    bom: "off",
    summary: false,
    outPath: null,
    zipPath: null
  };
  const positionals = [];

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (!arg.startsWith("--")) {
      positionals.push(arg);
      continue;
    }

    const flagHandler = FLAG_OPTIONS[arg];
    if (flagHandler) {
      flagHandler(options);
      continue;
    }

    const optionDefinition = valueOptions[arg];
    if (optionDefinition) {
      const value = argv[index + 1];
      if (!value) {
        throw new Error(`Missing value for ${arg}`);
      }
      index += 1;
      const normalizedValue = typeof optionDefinition.validate === "function"
        ? optionDefinition.validate(value)
        : value;
      optionDefinition.apply(options, normalizedValue);
      continue;
    }

    throw new Error(`Unknown option: ${arg}`);
  }

  if (positionals.length === 0) {
    options.inputPath = null;
  } else if (positionals.length === 1) {
    [options.inputPath] = positionals;
  } else {
    throw new Error("Specify exactly one input workbook.");
  }
  return options;
}

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function writeBinaryFile(outputPath, content) {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, content);
}

function formatWorkbookError(inputPath, stage, error) {
  const message = error instanceof Error ? error.message : String(error);
  return `[${path.basename(inputPath)}] ${stage}: ${message}`;
}

function printWorkbookSummary(api, workbookName, files) {
  console.log(`[workbook] ${workbookName}`);
  for (const file of files) {
    console.log(api.createSummaryText(file));
    console.log("");
  }
}

async function main() {
  const api = loadXlsx2mdNodeApi();
  const options = parseArgs(process.argv.slice(2), api.markdownOptions);
  if (options.help || !options.inputPath) {
    printHelp();
    process.exit(options.help ? 0 : 1);
  }

  const inputPath = path.resolve(options.inputPath);

  try {
    if (options.encoding === "shift_jis" && options.bom === "on") {
      throw new Error(formatWorkbookError(inputPath, "option validation failed", "BOM cannot be enabled for shift_jis."));
    }

    let inputBytes;
    try {
      inputBytes = await fs.readFile(inputPath);
    } catch (error) {
      throw new Error(formatWorkbookError(inputPath, "read failed", error));
    }

    let workbook;
    try {
      workbook = await api.parseWorkbook(toArrayBuffer(inputBytes), path.basename(inputPath));
    } catch (error) {
      throw new Error(formatWorkbookError(inputPath, "parse failed", error));
    }

    let files;
    try {
      files = api.convertWorkbookToMarkdownFiles(workbook, {
        treatFirstRowAsHeader: options.treatFirstRowAsHeader,
        trimText: options.trimText,
        removeEmptyRows: options.removeEmptyRows,
        removeEmptyColumns: options.removeEmptyColumns,
        includeShapeDetails: options.includeShapeDetails,
        outputMode: options.outputMode,
        formattingMode: options.formattingMode,
        tableDetectionMode: options.tableDetectionMode
      });
    } catch (error) {
      throw new Error(formatWorkbookError(inputPath, "convert failed", error));
    }

    if (options.summary) {
      printWorkbookSummary(api, path.basename(inputPath), files);
    }

    const combined = api.createCombinedMarkdownExportPayload(workbook, files, {
      encoding: options.encoding,
      bom: options.bom
    });

    if (options.zipPath) {
      try {
        const zipBytes = api.createWorkbookExportArchive(workbook, files, {
          encoding: options.encoding,
          bom: options.bom
        });
        await writeBinaryFile(path.resolve(options.zipPath), zipBytes);
      } catch (error) {
        throw new Error(formatWorkbookError(inputPath, "zip write failed", error));
      }
    }

    if (!options.zipPath || options.outPath) {
      const markdownOutputPath = options.outPath
        ? path.resolve(options.outPath)
        : path.resolve(combined.fileName);
      try {
        await writeBinaryFile(markdownOutputPath, combined.data);
      } catch (error) {
        throw new Error(formatWorkbookError(inputPath, "markdown write failed", error));
      }
    }
  } catch (error) {
    if (error instanceof Error && error.message.startsWith(`[${path.basename(inputPath)}] `)) {
      throw error;
    }
    throw new Error(formatWorkbookError(inputPath, "failed", error));
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
