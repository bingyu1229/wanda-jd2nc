/**
 * Convert Kingdee-style HTML table export (.htm) to CSV without changing cell text.
 *
 * - Reads GB2312/GBK HTML (default encoding gb18030).
 * - Writes every field quoted (Excel-style); no numeric parsing (preserves 2023.10.10, etc.).
 * - Streams line-by-line for large files.
 * - Drops the first table row (Kingdee placeholder headers 第1列…第N列).
 * - Drops column C (0-based index 2, 助记码) on every output row.
 * - Column A (instruction section 2): (2.1–2.2) dotted alphanumeric codes shaped to 4.(3)* — first segment
 *   width 4 (right-pad; overflow carries into next segment); each later dot-separated segment
 *   becomes one or more 3-character groups (per-segment; last chunk right-padded with zeroes). Pure 4-character
 *   codes stay unchanged (no trailing groups). (2.4) remove all "." so "1234.567.890" →
 *   "1234567890".
 *
 * Also writes 科目导入.csv next to the output CSV:
 * - A1 is the Kingdee import head marker; later A cells are blank.
 * - Source CSV columns A-D are copied to columns B-E.
 * - Column F is 有效期; data rows use 2000-01.
 *
 * Run: npx tsx htm_table_to_csv.ts <input.htm> [-o out.csv] [--encoding gb18030] [--utf8-bom] [--excel-text-cols 0,1] [--no-normalize-col-a]
 */

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

/**
 * Shape column A to 4.(3)*: first segment normalized to 4 characters (right-pad); overflow
 * digits from a long first segment are prefixed onto the next dot-separated segment.
 * Each segment after the first is expanded independently into 3-character chunks (left to right,
 * last chunk right-padded). No tail groups if there are no segments after the first
 * (e.g. "1001" stays "1001"). Example: "1151.13.01" → "1151.130.010" (not "1151.130.100");
 * "2131.003.HA11" → "2131.003.HA1.100".
 */
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
  const previousCount = duplicateBaseCounts.get(dotted) ?? 0;
  duplicateBaseCounts.set(dotted, previousCount + 1);

  if (previousCount === 0 && !usedDottedCodes.has(dotted)) {
    usedDottedCodes.add(dotted);
    return dotted;
  }

  let suffix = Math.max(previousCount, 1);
  while (true) {
    const candidate = `${dotted}.${String(suffix).padStart(3, "0")}`;
    if (!usedDottedCodes.has(candidate)) {
      usedDottedCodes.add(candidate);
      return candidate;
    }
    suffix++;
  }
}

function appendCodeToDuplicateNames(rows: string[][]): void {
  const nameCounts = new Map<string, number>();
  for (let i = 1; i < rows.length; i++) {
    const name = rows[i]?.[1] ?? "";
    if (!name) continue;
    nameCounts.set(name, (nameCounts.get(name) ?? 0) + 1);
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const code = row[0] ?? "";
    const name = row[1] ?? "";
    if (name && (nameCounts.get(name) ?? 0) > 1) {
      row[1] = `${name}${code}`;
    }
  }
}

/** Kingdee export: row 1 is 第1列… placeholders; column index 2 is 助记码 (omitted from CSV). */
const DROP_COL_INDEX = 2;
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
    output?: string;
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
    if (a === "-o" || a === "--output") {
      out.output = argv[++i];
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
  console.error(`Usage: npx tsx htm_table_to_csv.ts <input.htm> [options]

Options:
  -o, --output <file>     Output CSV (default: input basename + .csv)
  --encoding <name>       Input encoding (default: gb18030)
  --utf8-bom              Write UTF-8 with BOM for Excel on Windows
  --no-normalize-col-a    Skip column A normalization (instruction section 2, 2.1–2.4)
  --excel-text-cols <n>   Comma-separated 0-based column indexes as Excel text (="...")

Also writes ${SUBJECT_IMPORT_FILE_NAME} next to the output CSV.
`);
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args.input) {
    printHelp();
    return args.help ? 0 : 1;
  }

  const inPath = path.resolve(args.input);
  const outPath = args.output
    ? path.resolve(args.output)
    : inPath.replace(/\.htm$/i, ".csv");
  const subjectImportPath = path.join(
    path.dirname(outPath),
    SUBJECT_IMPORT_FILE_NAME,
  );

  if (!iconv.encodingExists(args.encoding)) {
    console.error("Unknown encoding:", args.encoding);
    return 2;
  }

  const readStream = fs.createReadStream(inPath);
  const decodeStream = iconv.decodeStream(args.encoding);
  const lineStream = readline.createInterface({
    input: readStream.pipe(decodeStream),
    crlfDelay: Number.POSITIVE_INFINITY,
  });

  const rows: string[][] = [];
  const usedDottedCodes = new Set<string>();
  const duplicateBaseCounts = new Map<string, number>();

  let htmTableRowIndex = 0;
  for await (const row of htmTableRows(lineStream)) {
    htmTableRowIndex++;
    if (htmTableRowIndex === 1) continue;

    let outRow = dropMnemonicColumn(row);
    const isOutputHeader = rows.length === 0;
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
          outRow = outRow.slice();
          outRow[0] = normalized;
        }
      }
    }
    rows.push(outRow);
    if (rows.length % 50000 === 0) {
      console.error(rows.length, "rows...");
    }
  }

  appendCodeToDuplicateNames(rows);

  const writeStream = fs.createWriteStream(outPath, { encoding: "utf8" });
  const subjectImportWriteStream = fs.createWriteStream(subjectImportPath, {
    encoding: "utf8",
  });
  if (args.utf8Bom) {
    writeStream.write("\uFEFF");
    subjectImportWriteStream.write("\uFEFF");
  }

  for (let i = 0; i < rows.length; i++) {
    let outRow = rows[i] ?? [];
    const isOutputHeader = i === 0;
    if (!isOutputHeader && args.excelTextCols.size > 0) {
      outRow = outRow.map((v, i) =>
        args.excelTextCols.has(i) ? excelTextFormula(v) : v,
      );
    }
    const lineOut = csvQuoteAllRow(outRow);
    if (!writeStream.write(lineOut)) {
      await once(writeStream, "drain");
    }
    const subjectImportLineOut = csvQuoteAllRow(
      subjectImportRowFromCsvRow(outRow, isOutputHeader),
    );
    if (!subjectImportWriteStream.write(subjectImportLineOut)) {
      await once(subjectImportWriteStream, "drain");
    }
  }

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

  console.error("Wrote", rows.length, "rows to", outPath);
  console.error("Wrote", rows.length, "rows to", subjectImportPath);
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

function subjectImportRowFromCsvRow(row: string[], isHeader: boolean): string[] {
  const copied = row.slice(0, 4);
  while (copied.length < 4) copied.push("");
  return [
    isHeader ? SUBJECT_IMPORT_HEAD : "",
    ...copied,
    isHeader ? SUBJECT_IMPORT_PERIOD_HEADER : SUBJECT_IMPORT_PERIOD_VALUE,
  ];
}
