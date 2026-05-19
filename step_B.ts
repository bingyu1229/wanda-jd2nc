import fs from "node:fs";
import path from "node:path";
import readline from "node:readline";
import { once } from "node:events";
import { pathToFileURL } from "node:url";
import iconv from "iconv-lite";
import { pinyin } from "pinyin-pro";

const CELL_LINE = /^\s*<T[DH]\b[^>]*>(?<body>.*?)<\/T[DH]>\s*$/i;

const DEFAULT_INPUT_FILE_NAME = "会计分类序时簿.htm";
const OUTPUT_CSV_FILE_NAME = "b.csv";
const FREEVALUE_FILE_NAME = "GL_FREEVALUE.csv";
const NAME_SQL_FILE_NAME = "name_sql.txt";
const ENTRY_NO_HEADER = "分录号";

const COL_DATE = 0;
const COL_PERIOD = 1;
const COL_VOUCHER_NO = 2;
const COL_SUBJECT_CODE = 4;
const COL_SUBJECT_NAME = 5;
const COL_ORIGINAL_AMOUNT = 8;
const COL_DEBIT = 9;
const COL_CREDIT = 10;
const COL_PREPARED_BY = 11;
const COL_AUDITED_BY = 12;
const COL_POSTED_BY = 13;
const COL_BUSINESS_DATE = 21;

const FREEVALUE_HEADER = [
  "ASSINDEX",
  "CHECKCOUNT",
  "CHECKTYPE",
  "CHECKVALUE",
  "DR",
  "FREE1",
  "FREE2",
  "FREE3",
  "FREEVALUEID",
  "PK_FREEVALUE",
  "TS",
  "VALUECODE",
  "VALUENAME",
];

const FREEVALUE_CHECKTYPE = "0001A9100000000JCKUS";
const FREEVALUE_TS = "2026/3/5 16:26";
const FREEVALUE_CHECKVALUE_UUID_START = 150000001;
const FREEVALUE_ID_UUID_PREFIX = "15010000000";
const FREEVALUE_ID_UUID_START = parseInt("391", 36);
const USER_CODE_SUFFIX = "(jindie)";

const NAME_PINYIN_OVERRIDES = new Map<string, string>([
  // Keep this spelling aligned with the requested sample output.
  ["张春光", "zhangchunghuang"],
]);

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

function shouldNormalizeCode(s: string): boolean {
  const t = s.trim();
  return !!t && /^[A-Za-z0-9.]+$/.test(t);
}

function stripIntegerDecimalArtifact(s: string): string {
  return /^\d+\.0$/.test(s) ? s.slice(0, -2) : s;
}

function normalizeDottedCode(raw: string): string {
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

  const { g0, overflow } = head4(parts[0] ?? "");
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

  return tailGroups.length ? [g0, ...tailGroups].join(".") : g0;
}

function flattenCode(dotted: string): string {
  return dotted.replace(/\./g, "");
}

function csvQuoteField(value: string): string {
  const escaped = value.replace(/"/g, '""');
  return `"${escaped}"`;
}

function csvQuoteAllRow(fields: string[]): string {
  return fields.map(csvQuoteField).join(",") + "\n";
}

function csvRow(fields: string[]): string {
  return (
    fields
      .map((value) =>
        /[",\r\n]/.test(value) ? csvQuoteField(value) : value,
      )
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
    if (!m?.groups) continue;
    row.push(cellTextFromBody(m.groups.body ?? ""));
  }
  if (row.length) yield row;
}

function parseArgs(argv: string[]) {
  const out: {
    input?: string;
    encoding: string;
    utf8Bom: boolean;
    excelTextCols: Set<number>;
    normalizeColE: boolean;
    help: boolean;
  } = {
    encoding: "gb18030",
    utf8Bom: false,
    excelTextCols: new Set(),
    normalizeColE: true,
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
    if (a === "--no-normalize-col-e") {
      out.normalizeColE = false;
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
  console.error(`Usage: npx tsx step_B.ts [input.htm] [options]

Options:
  input.htm               Optional single input file. When omitted, recursively processes every ${DEFAULT_INPUT_FILE_NAME} under the current folder.
  --encoding <name>       Input encoding (default: gb18030)
  --utf8-bom              Write UTF-8 with BOM for Excel on Windows
  --no-normalize-col-e    Skip column E account code normalization
  --excel-text-cols <n>   Comma-separated 0-based column indexes as Excel text (="...")

Writes ${OUTPUT_CSV_FILE_NAME}, ${FREEVALUE_FILE_NAME}, and ${NAME_SQL_FILE_NAME} next to each input file.
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

function fillDownColumns(rows: string[][], colIndexes: number[]): void {
  const previous = new Map<number, string>();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    for (const col of colIndexes) {
      const value = row[col] ?? "";
      if (value) previous.set(col, value);
      else if (previous.has(col)) row[col] = previous.get(col) ?? "";
    }
  }
}

function fillEntryNumbers(rows: string[][]): void {
  if (rows.length === 0) return;
  rows[0].push(ENTRY_NO_HEADER);

  let previousKey = "";
  let entryNo = 0;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const key = [
      row[COL_DATE] ?? "",
      row[COL_PERIOD] ?? "",
      row[COL_VOUCHER_NO] ?? "",
    ].join("\u0001");
    if (key !== previousKey) {
      previousKey = key;
      entryNo = 1;
    } else {
      entryNo++;
    }
    row.push(String(entryNo));
  }
}

function normalizeAmount(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) return value;

  const withoutTextMarker = trimmed.startsWith("'") ? trimmed.slice(1) : trimmed;
  const normalizedNumberText = withoutTextMarker.replace(/,/g, "");
  const amount = Number(normalizedNumberText);
  if (!Number.isFinite(amount)) return withoutTextMarker;

  return Number.isInteger(amount) ? String(amount) : amount.toFixed(2);
}

function cleanAmountColumns(rows: string[][]): void {
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    for (const col of [COL_ORIGINAL_AMOUNT, COL_DEBIT, COL_CREDIT]) {
      row[col] = normalizeAmount(row[col] ?? "");
    }
  }
}

function cleanSubjectNames(rows: string[][]): void {
  for (let i = 1; i < rows.length; i++) {
    const value = rows[i][COL_SUBJECT_NAME] ?? "";
    rows[i][COL_SUBJECT_NAME] = value.replace(/b/g, "").replace(/,/g, "");
  }
}

function auxiliaryValueFromSubjectName(subjectName: string): string {
  const marker = " - ";
  const markerIndex = subjectName.indexOf(marker);
  if (markerIndex < 0) return subjectName;
  return subjectName.slice(markerIndex + marker.length).trim();
}

function hasAuxiliaryValue(subjectName: string): boolean {
  return subjectName.includes(" - ");
}

function freeValueUuidForIndex(index: number): string {
  return (
    FREEVALUE_ID_UUID_PREFIX +
    (FREEVALUE_ID_UUID_START + index)
      .toString(36)
      .toUpperCase()
      .padStart(3, "0")
  );
}

function freeValueIdForIndex(index: number): string {
  return `1774A${freeValueUuidForIndex(index)}F`;
}

function freeValueRows(values: string[]): string[][] {
  return values.map((value, index) => {
    const ordinal = index + 1;
    const valueCode = String(ordinal).padStart(5, "0");
    const checkValueUuid = String(FREEVALUE_CHECKVALUE_UUID_START + index);
    const freeValueUuid = freeValueUuidForIndex(index);
    return [
      "0",
      "1",
      FREEVALUE_CHECKTYPE,
      `0001A92JDT${checkValueUuid}U`,
      "0",
      "",
      "",
      "",
      `1774A${freeValueUuid}F`,
      `1774A${freeValueUuid}P`,
      FREEVALUE_TS,
      valueCode,
      value,
    ];
  });
}

function auxiliaryFreeValues(rows: string[][]): string[] {
  const values: string[] = [];
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    const subjectName = rows[rowIndex][COL_SUBJECT_NAME] ?? "";
    if (!hasAuxiliaryValue(subjectName)) continue;
    values.push(auxiliaryValueFromSubjectName(subjectName));
  }
  return values;
}

function collectUserNames(rows: string[][]): string[] {
  const names = new Set<string>();
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    for (const col of [COL_PREPARED_BY, COL_AUDITED_BY, COL_POSTED_BY]) {
      const name = (row[col] ?? "").trim();
      if (name) names.add(name);
    }
  }
  return [...names];
}

function userCodeForName(name: string): string {
  const override = NAME_PINYIN_OVERRIDES.get(name);
  if (override) return `${override}${USER_CODE_SUFFIX}`;

  const code = pinyin(name, {
    toneType: "none",
    type: "array",
    nonZh: "consecutive",
  })
    .join("")
    .replace(/\s+/g, "")
    .toLowerCase();

  return `${code}${USER_CODE_SUFFIX}`;
}

function sqlQuotedList(values: string[]): string[] {
  const lines = ["("];
  values.forEach((value, index) => {
    const escaped = value.replace(/'/g, "''");
    const suffix = index === values.length - 1 ? "" : ",";
    lines.push(`'${escaped}'${suffix}`);
  });
  lines.push(")");
  return lines;
}

function nameSqlLines(names: string[]): string[] {
  const userCodes = names.map(userCodeForName);
  return [
    "USER_CODE",
    ...sqlQuotedList(userCodes),
    "",
    "USER_NAME",
    ...sqlQuotedList(names),
  ].map((line) => `${line}\n`);
}

async function writeLines(
  outPath: string,
  lines: Iterable<string>,
  utf8Bom: boolean,
): Promise<void> {
  const writeStream = fs.createWriteStream(outPath, { encoding: "utf8" });
  const streamDone = new Promise<void>((resolve, reject) => {
    writeStream.on("error", reject);
    writeStream.on("finish", resolve);
  });

  if (utf8Bom) writeStream.write("\uFEFF");
  for (const line of lines) {
    if (!writeStream.write(line)) await once(writeStream, "drain");
  }
  writeStream.end();
  await streamDone;
}

async function processInputFile(
  inPath: string,
  args: ReturnType<typeof parseArgs>,
): Promise<number> {
  const outPath = path.join(path.dirname(inPath), OUTPUT_CSV_FILE_NAME);
  const freeValuePath = path.join(path.dirname(inPath), FREEVALUE_FILE_NAME);
  const nameSqlPath = path.join(path.dirname(inPath), NAME_SQL_FILE_NAME);

  const readStream = fs.createReadStream(inPath);
  const decodeStream = iconv.decodeStream(args.encoding);
  const lineStream = readline.createInterface({
    input: readStream.pipe(decodeStream),
    crlfDelay: Number.POSITIVE_INFINITY,
  });

  const rows: string[][] = [];

  let htmTableRowIndex = 0;
  for await (const row of htmTableRows(lineStream)) {
    htmTableRowIndex++;
    if (htmTableRowIndex === 1) continue;

    let outRow = row.slice();
    const isOutputHeader = rows.length === 0;

    if (!isOutputHeader && args.normalizeColE && outRow.length > COL_SUBJECT_CODE) {
      const code = outRow[COL_SUBJECT_CODE] ?? "";
      if (shouldNormalizeCode(code)) {
        const dotted = normalizeDottedCode(code);
        outRow[COL_SUBJECT_CODE] = flattenCode(dotted);
      }
    }
    rows.push(outRow);
  }

  fillDownColumns(rows, [COL_DATE, COL_PERIOD, COL_VOUCHER_NO, COL_BUSINESS_DATE]);
  fillEntryNumbers(rows);
  cleanAmountColumns(rows);
  
  cleanSubjectNames(rows);

  const outputRows = rows.map((row, rowIndex) => {
    if (rowIndex === 0 || args.excelTextCols.size === 0) return row;
    return row.map((value, colIndex) =>
      args.excelTextCols.has(colIndex) ? excelTextFormula(value) : value,
    );
  });

  await writeLines(outPath, outputRows.map((row) => csvQuoteAllRow(row)), args.utf8Bom);

  const freeValues = auxiliaryFreeValues(rows);

  await writeLines(
    freeValuePath,
    [csvRow(FREEVALUE_HEADER), ...freeValueRows(freeValues).map((row) => csvRow(row))],
    args.utf8Bom,
  );

  const userNames = collectUserNames(rows);
  await writeLines(nameSqlPath, nameSqlLines(userNames), args.utf8Bom);

  console.error("Wrote", rows.length, "rows to", outPath);
  console.error("Wrote", freeValues.length + 1, "rows to", freeValuePath);
  console.error("Wrote", userNames.length, "names to", nameSqlPath);
  return rows.length;
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
