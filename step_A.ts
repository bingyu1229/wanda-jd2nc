import fs from "node:fs";
import path from "node:path";
import readline from "node:readline";
import { once } from "node:events";
import { pathToFileURL } from "node:url";
import iconv from "iconv-lite";

const CELL_LINE = /^\s*<T[DH]\b[^>]*>(?<body>.*?)<\/T[DH]>\s*$/i;

function htmlUnescape(s: string): string {
  return s
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n: string) => String.fromCodePoint(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/gi, (_, h: string) =>
      String.fromCodePoint(parseInt(h, 16)),
    );
}

function cellTextFromBody(body: string): string {
  let t = htmlUnescape(body);
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, " ");
  return t.trim();
}

function excelTextFormula(value: string): string {
  const escaped = value.replace(/"/g, '""');
  return `="${escaped}"`;
}

/** Dotted alphanumeric account codes only; excludes text headers and non-code text. */
export function shouldNormalizeColA(s: string): boolean {
  const t = s.trim();
  if (!t || !/^[A-Za-z0-9.]+$/.test(t)) return false;
  return true;
}

function stripIntegerDecimalArtifact(s: string): string {
  return /^\d+\.0$/.test(s) ? s.slice(0, -2) : s;
}

function normalizeColADottedCode(raw: string): string {
  const t = stripIntegerDecimalArtifact(raw.trim()).replace(/\.+$/, "");
  if (!t) return raw;
  const parts = t.split(".").filter((p) => p.length > 0);
  if (parts.length === 0) return raw;

  function head4(seg: string): { g0: string; overflow: string } {
    if (!seg) return { g0: "0000", overflow: "" };
    if (seg.length <= 4)
      return { g0: seg + "0".repeat(4 - seg.length), overflow: "" };
    return { g0: seg.slice(0, 4), overflow: seg.slice(4) };
  }

  const p0 = parts[0] ?? "";
  const { g0, overflow } = head4(p0);
  const restParts = parts.slice(1);
  if (overflow) {
    if (restParts.length) restParts[0] = overflow + restParts[0];
    else restParts.push(overflow);
  }

  const tailGroups: string[] = [];
  for (const seg of restParts) {
    let s = seg;
    while (s.length > 0) {
      const chunk = s.slice(0, 3);
      s = s.slice(3);
      tailGroups.push(
        chunk.length < 3 ? chunk + "0".repeat(3 - chunk.length) : chunk,
      );
    }
  }

  if (tailGroups.length === 0) return g0;
  return [g0, ...tailGroups].join(".");
}

/** Instruction 2.4: drop dot separators after 4.(3)* shaping (e.g. "1234.567.890" → "1234567890"). */
function flattenColAAccountCode(dotted: string): string {
  return dotted.replace(/\./g, "");
}

export function normalizeColAAccountCode(raw: string): string {
  return flattenColAAccountCode(normalizeColADottedCode(raw));
}

function makeUniqueDottedCode(
  dotted: string,
  usedDottedCodes: Set<string>,
  duplicateBaseCounts: Map<string, number>,
): string {
  // 第一次出现：直接保留原值
  if (!usedDottedCodes.has(dotted)) {
    usedDottedCodes.add(dotted);
    duplicateBaseCounts.set(dotted, 0);
    return dotted;
  }

  let suffix = (duplicateBaseCounts.get(dotted) ?? 0) + 1;

  while (true) {
    const alphaSuffix = alphaSequence(suffix);

    const candidate =
      dotted.length > alphaSuffix.length
        ? dotted.slice(0, -alphaSuffix.length) + alphaSuffix
        : alphaSuffix;

    if (!usedDottedCodes.has(candidate)) {
      usedDottedCodes.add(candidate);
      duplicateBaseCounts.set(dotted, suffix);
      return candidate;
    }

    suffix++;
  }
}

function alphaSequence(n: number): string {
  let s = "";
  let current = n;
  while (current > 0) {
    current--;
    s = String.fromCharCode(97 + (current % 26)) + s;
    current = Math.floor(current / 26);
  }
  return s;
}

/** Kingdee export: row 1 is 第1列… placeholders; column index 2 is 助记码 (omitted from CSV). */
const DROP_COL_INDEX = 2;
const DEFAULT_INPUT_FILE_NAME = "科目.htm";
const SUBJECT_IMPORT_FILE_NAME = "科目导入.csv";
const SUBJECT_IMPORT_HEAD =
  "bd_accsubj_$head,subjcode,subjname,pk_subjtype,balanorient,period,outflag";
const SUBJECT_IMPORT_PERIOD_HEADER = "有效期";
const SUBJECT_IMPORT_PERIOD_VALUE = "2000-01";

function dropMnemonicColumn(row: string[]): string[] {
  return row.filter((_, i) => i !== DROP_COL_INDEX);
}

/** RFC 4180 / Excel: comma, all fields double-quoted */
function csvQuoteAllRow(fields: string[]): string {
  return (
    fields
      .map((f) => {
        const s = f.replace(/"/g, '""');
        return `"${s}"`;
      })
      .join(",") + "\n"
  );
}

async function* htmTableRows(
  lines: AsyncIterable<string>,
): AsyncGenerator<string[]> {
  let row: string[] = [];
  let inTable = false;

  for await (const line of lines) {
    const ul = line.toUpperCase();
    if (ul.includes("<TABLE")) {
      inTable = true;
      continue;
    }
    if (!inTable) continue;
    if (ul.includes("</TABLE>")) {
      if (row.length) yield row;
      break;
    }
    if (ul.includes("</TR>")) {
      if (row.length) yield row;
      row = [];
      continue;
    }
    if (ul.includes("<TR")) {
      if (row.length) yield row;
      row = [];
      continue;
    }
    const m = CELL_LINE.exec(line);
    if (!m?.groups?.body) continue;
    row.push(cellTextFromBody(m.groups.body));
  }
  if (row.length) yield row;
}

function parseArgs(argv: string[]) {
  const out: {
    input?: string;
    encoding: string;
    utf8Bom: boolean;
    excelTextCols: Set<number>;
    normalizeColA: boolean;
    help: boolean;
  } = {
    encoding: "gb18030",
    utf8Bom: false,
    excelTextCols: new Set(),
    normalizeColA: true,
    help: false,
  };

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "-h" || a === "--help") {
      out.help = true;
      continue;
    }
    if (a === "--encoding") {
      out.encoding = argv[++i] ?? out.encoding;
      continue;
    }
    if (a === "--utf8-bom") {
      out.utf8Bom = true;
      continue;
    }
    if (a === "--no-normalize-col-a") {
      out.normalizeColA = false;
      continue;
    }
    if (a === "--excel-text-cols") {
      const list = argv[++i] ?? "";
      for (const part of list.split(",")) {
        const p = part.trim();
        if (p) out.excelTextCols.add(parseInt(p, 10));
      }
      continue;
    }
    if (!a.startsWith("-")) {
      out.input = a;
      continue;
    }
    console.error("Unknown argument:", a);
    process.exit(2);
  }
  return out;
}

function printHelp(): void {
  console.error(`Usage: npx tsx step_A.ts [input.htm] [options]

Options:
  input.htm               Optional single input file. When omitted, recursively processes every ${DEFAULT_INPUT_FILE_NAME} under the current folder.
  --encoding <name>       Input encoding (default: gb18030)
  --utf8-bom              Write UTF-8 with BOM for Excel on Windows
  --no-normalize-col-a    Skip column A normalization (instruction section 2, 2.1–2.4)
  --excel-text-cols <n>   Comma-separated 0-based column indexes as Excel text (="...")

Writes a.csv and ${SUBJECT_IMPORT_FILE_NAME} next to each input file.
`);
}

async function findDefaultInputFiles(rootDir: string): Promise<string[]> {
  const found: string[] = [];

  async function walk(dir: string): Promise<void> {
    const entries = await fs.promises.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.name === ".git" || entry.name === "node_modules") continue;

      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        await walk(fullPath);
      } else if (entry.isFile() && entry.name === DEFAULT_INPUT_FILE_NAME) {
        found.push(fullPath);
      }
    }
  }

  await walk(rootDir);
  return found.sort((a, b) => a.localeCompare(b));
}

/**
 * Phase 1: Lightweight scan to count name frequencies.
 * Only collects {name -> count}, not full rows. Memory O(unique names).
 */
async function scanNameFrequencies(
  inPath: string,
  encoding: string,
): Promise<Map<string, number>> {
  const readStream = fs.createReadStream(inPath);
  const decodeStream = iconv.decodeStream(encoding);
  const lineStream = readline.createInterface({
    input: readStream.pipe(decodeStream),
    crlfDelay: Number.POSITIVE_INFINITY,
  });

  const nameFreq = new Map<string, number>();
  let htmTableRowIndex = 0;

  for await (const row of htmTableRows(lineStream)) {
    htmTableRowIndex++;
    if (htmTableRowIndex === 1) continue; // Skip HTML header row

    const outRow = dropMnemonicColumn(row);
    const name = outRow[1] ?? "";
    if (name) {
      nameFreq.set(name, (nameFreq.get(name) ?? 0) + 1);
    }
  }

  return nameFreq;
}

function subjectImportRowFromCsvRow(row: string[], isHeader: boolean): string[] {
  const copied = row.slice(0, 4);
  while (copied.length < 4) copied.push("");
  return [
    isHeader ? SUBJECT_IMPORT_HEAD : "",
    ...copied,
    isHeader ? SUBJECT_IMPORT_PERIOD_HEADER : SUBJECT_IMPORT_PERIOD_VALUE,
  ];
}

/**
 * Phase 2: Stream processing — read, transform, write immediately.
 * No full table stored in memory.
 */
async function processInputFileStreaming(
  inPath: string,
  args: ReturnType<typeof parseArgs>,
  nameFreq: Map<string, number>,
): Promise<number> {
  const outPath = path.join(path.dirname(inPath), "a.csv");
  const subjectImportPath = path.join(
    path.dirname(outPath),
    SUBJECT_IMPORT_FILE_NAME,
  );

  const readStream = fs.createReadStream(inPath);
  const decodeStream = iconv.decodeStream(args.encoding);
  const lineStream = readline.createInterface({
    input: readStream.pipe(decodeStream),
    crlfDelay: Number.POSITIVE_INFINITY,
  });

  const writeStream = fs.createWriteStream(outPath, { encoding: "utf8" });
  const subjectImportWriteStream = fs.createWriteStream(subjectImportPath, {
    encoding: "utf8",
  });

  if (args.utf8Bom) {
    writeStream.write("\uFEFF");
    subjectImportWriteStream.write("\uFEFF");
  }

  // State for code deduplication (must persist across rows)
  const usedDottedCodes = new Set<string>();
  const duplicateBaseCounts = new Map<string, number>();

  let rowCount = 0;
  let dataRowCount = 0;

  for await (const row of htmTableRows(lineStream)) {
    rowCount++;
    if (rowCount === 1) continue; // Skip HTML placeholder header

    let outRow = dropMnemonicColumn(row);
    const isOutputHeader = dataRowCount === 0;
    const originalCode = outRow[0] ?? "";

    // 2.1: Append original code as new last column
    outRow = outRow.slice();
    outRow.push(isOutputHeader ? "原科目代码" : originalCode);

    // 2.2–2.4: Normalize column A (account code)
    if (!isOutputHeader && args.normalizeColA && outRow.length > 0) {
      const a = outRow[0] ?? "";
      if (shouldNormalizeColA(a)) {
        const dotted = normalizeColADottedCode(a);
        const uniqueDotted = makeUniqueDottedCode(
          dotted,
          usedDottedCodes,
          duplicateBaseCounts,
        );
        const normalized = flattenColAAccountCode(uniqueDotted);
        if (normalized !== a) {
          outRow[0] = normalized;
        }
      }
    }

    // 3: Append code to duplicate names (using pre-scanned frequency map)
    if (!isOutputHeader) {
      const name = outRow[1] ?? "";
      if (name && (nameFreq.get(name) ?? 0) > 1) {
        outRow[1] = `${name}${outRow[0]}`;
      }
    }

    // Apply Excel text formula if requested
    if (!isOutputHeader && args.excelTextCols.size > 0) {
      outRow = outRow.map((v, i) =>
        args.excelTextCols.has(i) ? excelTextFormula(v) : v,
      );
    }

    // Write to a.csv
    const lineOut = csvQuoteAllRow(outRow);
    if (!writeStream.write(lineOut)) {
      await once(writeStream, "drain");
    }

    // Write to 科目导入.csv
    const subjectImportLineOut = csvQuoteAllRow(
      subjectImportRowFromCsvRow(outRow, isOutputHeader),
    );
    if (!subjectImportWriteStream.write(subjectImportLineOut)) {
      await once(subjectImportWriteStream, "drain");
    }

    dataRowCount++;
    if (dataRowCount % 50000 === 0) {
      console.error(dataRowCount, "rows processed...");
    }
  }

  // Close streams
  await new Promise<void>((resolve, reject) => {
    writeStream.end((err: NodeJS.ErrnoException | null) =>
      err ? reject(err) : resolve(),
    );
  });
  await new Promise<void>((resolve, reject) => {
    subjectImportWriteStream.end((err: NodeJS.ErrnoException | null) =>
      err ? reject(err) : resolve(),
    );
  });

  console.error("Wrote", dataRowCount, "rows to", outPath);
  console.error("Wrote", dataRowCount, "rows to", subjectImportPath);
  return dataRowCount;
}

async function processInputFile(
  inPath: string,
  args: ReturnType<typeof parseArgs>,
): Promise<number> {
  // Phase 1: Scan name frequencies (lightweight, memory-efficient)
  const nameFreq = await scanNameFrequencies(inPath, args.encoding);

  // Phase 2: Stream processing with the frequency map
  return processInputFileStreaming(inPath, args, nameFreq);
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return 0;
  }

  if (!iconv.encodingExists(args.encoding)) {
    console.error("Unknown encoding:", args.encoding);
    return 2;
  }

  const inputFiles = args.input
    ? [path.resolve(args.input)]
    : await findDefaultInputFiles(process.cwd());

  if (inputFiles.length === 0) {
    console.error(
      `No ${DEFAULT_INPUT_FILE_NAME} files found under ${process.cwd()}`,
    );
    return 1;
  }

  let totalRows = 0;
  for (const inPath of inputFiles) {
    console.error("Processing", inPath);
    totalRows += await processInputFile(inPath, args);
  }

  console.error(
    "Processed",
    inputFiles.length,
    "folder(s),",
    totalRows,
    "CSV row(s).",
  );
  return 0;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().then(
    (code) => process.exit(code),
    (err) => {
      console.error(err);
      process.exit(1);
    },
  );
}
