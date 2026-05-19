import fs from "node:fs";
import path from "node:path";
import { once } from "node:events";
import { pathToFileURL } from "node:url";

const INPUT_B_FILE_NAME = "b.csv";
const INPUT_FREEVALUE_FILE_NAME = "GL_FREEVALUE.csv";
const INPUT_PK_FILE_NAME = "pk.csv";
const OUTPUT_DETAIL_FILE_NAME = "GL_DETAIL.csv";

const COL_DATE = 0;
const COL_PERIOD = 1;
const COL_SUBJECT_CODE = 4;
const COL_SUBJECT_NAME = 5;
const COL_DEBIT = 9;
const COL_CREDIT = 10;
const COL_SUMMARY = 3;

const DETAIL_HEADER = [
  "ASSID",
  "BANKACCOUNT",
  "CHECKDATE",
  "CHECKNO",
  "CHECKSTYLE",
  "CONTRASTFLAG",
  "CONVERTFLAG",
  "CREDITAMOUNT",
  "CREDITQUANTITY",
  "DEBITAMOUNT",
  "DEBITQUANTITY",
  "DETAILINDEX",
  "DR",
  "ERRMESSAGE",
  "EXCRATE1",
  "EXCRATE2",
  "EXPLANATION",
  "FRACCREDITAMOUNT",
  "FRACDEBITAMOUNT",
  "FREE1",
  "FREE2",
  "FREE3",
  "FREE4",
  "FREE5",
  "LOCALCREDITAMOUNT",
  "LOCALDEBITAMOUNT",
  "MODIFYFLAG",
  "OPPOSITESUBJ",
  "PK_ACCSUBJ",
  "PK_CORP",
  "PK_CURRTYPE",
  "PK_DETAIL",
  "PK_GLBOOK",
  "PK_GLORG",
  "PK_GLORGBOOK",
  "PK_INNERCORP",
  "PK_INNERSOB",
  "PK_SOB",
  "PK_SOURCEPK",
  "PK_VOUCHER",
  "PRICE",
  "RECIEPTCLASS",
  "TS",
  "DIRECTION",
  "DISCARDFLAGV",
  "ERRMESSAGE2",
  "FREE6",
  "NOV",
  "PERIODV",
  "PK_MANAGERV",
  "PK_SYSTEMV",
  "PK_VOUCHERTYPEV",
  "PREPAREDDATEV",
  "SIGNDATEV",
  "VOUCHERKINDV",
  "YEARV",
  "BUSIRECONNO",
  "ERRMESSAGEH",
  "FREE10",
  "FREE11",
  "FREE7",
  "FREE8",
  "FREE9",
  "ISDIFFLAG",
  "PK_OFFERDETAIL",
  "PK_OTHERCORP",
  "PK_OTHERORGBOOK",
];

const DETAIL_PK_DETAIL_UUID_START = 15020000000001;
const DETAIL_PK_VOUCHER_UUID_START = 150000001;
const DETAIL_TS = "2026-03-11 9:00:00";
const DETAIL_PK_GLBOOK = "0001A9100000000JCNSC";
const DETAIL_PK_CURRTYPE = "00010000000000000001";
const DETAIL_MODIFYFLAG = "YYYYYYYYYYYYYYYY";
const DETAIL_PK_VOUCHERTYPEV = "0001DEFAULT000000001";

type SubjectPkRow = {
  pkAccsubjByCode: Map<string, string>;
  pkCorp: string;
  pkGlorg: string;
  pkGlorgbook: string;
};

function parseArgs(argv: string[]) {
  const out: {
    input?: string;
    utf8Bom: boolean;
    help: boolean;
  } = {
    utf8Bom: false,
    help: false,
  };

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "-h" || a === "--help") {
      out.help = true;
      continue;
    }
    if (a === "--utf8-bom") {
      out.utf8Bom = true;
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
  console.error(`Usage: npx tsx step_C.ts [folder-or-b.csv] [options]

Options:
  folder-or-b.csv    Optional folder or b.csv path. When omitted, recursively processes folders with ${INPUT_B_FILE_NAME}, ${INPUT_FREEVALUE_FILE_NAME}, and ${INPUT_PK_FILE_NAME}.
  --utf8-bom         Write UTF-8 with BOM for Excel on Windows

Reads ${INPUT_B_FILE_NAME}, ${INPUT_FREEVALUE_FILE_NAME}, and ${INPUT_PK_FILE_NAME}; writes ${OUTPUT_DETAIL_FILE_NAME} next to them.
`);
}

async function findInputFolders(rootDir: string): Promise<string[]> {
  const found: string[] = [];

  async function walk(dir: string): Promise<void> {
    const entries = await fs.promises.readdir(dir, { withFileTypes: true });
    const names = new Set(entries.map((entry) => entry.name));
    if (
      names.has(INPUT_B_FILE_NAME) &&
      names.has(INPUT_FREEVALUE_FILE_NAME) &&
      names.has(INPUT_PK_FILE_NAME)
    ) {
      found.push(dir);
    }

    for (const entry of entries) {
      if (entry.name === ".git" || entry.name === "node_modules") continue;
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) await walk(fullPath);
    }
  }

  await walk(rootDir);
  return found.sort((a, b) => a.localeCompare(b));
}

async function inputFoldersFromArg(input: string | undefined): Promise<string[]> {
  if (!input) return findInputFolders(process.cwd());

  const resolved = path.resolve(input);
  const stat = await fs.promises.stat(resolved);
  if (stat.isDirectory()) return [resolved];
  if (stat.isFile() && path.basename(resolved) === INPUT_B_FILE_NAME) {
    return [path.dirname(resolved)];
  }

  throw new Error(`Input must be a folder or ${INPUT_B_FILE_NAME}: ${input}`);
}

function parseCsv(content: string): string[][] {
  const text = content.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }
    if (ch === ",") {
      row.push(field);
      field = "";
      continue;
    }
    if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }
    if (ch === "\r") continue;
    field += ch;
  }

  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

async function readCsv(filePath: string): Promise<string[][]> {
  return parseCsv(await fs.promises.readFile(filePath, "utf8"));
}

function csvQuoteField(value: string): string {
  const escaped = value.replace(/"/g, '""');
  return `"${escaped}"`;
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

function hasAuxiliaryValue(subjectName: string): boolean {
  return subjectName.includes(" - ");
}

function auxiliaryValueFromSubjectName(subjectName: string): string {
  const marker = " - ";
  const markerIndex = subjectName.indexOf(marker);
  if (markerIndex < 0) return subjectName;
  return subjectName.slice(markerIndex + marker.length).trim();
}

function parsePeriod(period: string): { year: string; month: string } {
  const match = period.trim().match(/^(\d{4})\D+(\d{1,2})/);
  if (!match) return { year: "", month: "" };
  return {
    year: match[1],
    month: match[2].padStart(2, "0"),
  };
}

function formatVoucherDate(date: string): string {
  const match = date.trim().match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!match) return date;
  return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
}

function isZeroAmount(value: string): boolean {
  const amount = Number(value.replace(/,/g, "").trim() || "0");
  return Number.isFinite(amount) && amount === 0;
}

function freeValueIdColumnIndex(freeValueHeader: string[]): number {
  const index = freeValueHeader.indexOf("FREEVALUEID");
  if (index < 0) throw new Error(`${INPUT_FREEVALUE_FILE_NAME} is missing FREEVALUEID`);
  return index;
}

function valueNameColumnIndex(freeValueHeader: string[]): number {
  const index = freeValueHeader.indexOf("VALUENAME");
  if (index < 0) throw new Error(`${INPUT_FREEVALUE_FILE_NAME} is missing VALUENAME`);
  return index;
}

function requiredColumnIndex(header: string[], name: string, fileName: string): number {
  const index = header.indexOf(name);
  if (index < 0) throw new Error(`${fileName} is missing ${name}`);
  return index;
}

function subjectPkRowFromCsv(pkRows: string[][]): SubjectPkRow {
  const header = pkRows[0] ?? [];
  const subjCodeCol = requiredColumnIndex(header, "SUBJCODE", INPUT_PK_FILE_NAME);
  const pkAccsubjCol = requiredColumnIndex(header, "PK_ACCSUBJ", INPUT_PK_FILE_NAME);
  const pkCorpCol = requiredColumnIndex(header, "PK_CORP", INPUT_PK_FILE_NAME);
  const pkGlorgCol = requiredColumnIndex(header, "PK_GLORG", INPUT_PK_FILE_NAME);
  const pkGlorgbookCol = requiredColumnIndex(
    header,
    "PK_GLORGBOOK",
    INPUT_PK_FILE_NAME,
  );

  const firstDataRow = pkRows[1];
  if (!firstDataRow) throw new Error(`${INPUT_PK_FILE_NAME} has no data rows`);

  const pkAccsubjByCode = new Map<string, string>();
  for (let rowIndex = 1; rowIndex < pkRows.length; rowIndex++) {
    const row = pkRows[rowIndex];
    const subjCode = (row[subjCodeCol] ?? "").trim();
    if (!subjCode) continue;
    pkAccsubjByCode.set(subjCode, row[pkAccsubjCol] ?? "");
  }

  return {
    pkAccsubjByCode,
    pkCorp: firstDataRow[pkCorpCol] ?? "",
    pkGlorg: firstDataRow[pkGlorgCol] ?? "",
    pkGlorgbook: firstDataRow[pkGlorgbookCol] ?? "",
  };
}

function assIdsByDataRowIndex(
  bRows: string[][],
  freeValueRows: string[][],
): Map<number, string> {
  const map = new Map<number, string>();
  const freeValueIdCol = freeValueIdColumnIndex(freeValueRows[0] ?? []);
  const valueNameCol = valueNameColumnIndex(freeValueRows[0] ?? []);
  let freeValueDataIndex = 1;

  for (let bRowIndex = 1; bRowIndex < bRows.length; bRowIndex++) {
    const subjectName = bRows[bRowIndex][COL_SUBJECT_NAME] ?? "";
    if (!hasAuxiliaryValue(subjectName)) continue;

    const freeValueRow = freeValueRows[freeValueDataIndex];
    if (!freeValueRow) {
      throw new Error(
        `${INPUT_FREEVALUE_FILE_NAME} has fewer auxiliary rows than ${INPUT_B_FILE_NAME}`,
      );
    }

    const expectedValue = auxiliaryValueFromSubjectName(subjectName);
    const actualValue = freeValueRow[valueNameCol] ?? "";
    if (expectedValue !== actualValue) {
      console.error(
        `Warning: row ${bRowIndex + 1} auxiliary value differs from GL_FREEVALUE row ${freeValueDataIndex + 1}.`,
      );
    }

    map.set(bRowIndex - 1, freeValueRow[freeValueIdCol] ?? "");
    freeValueDataIndex++;
  }

  if (freeValueDataIndex < freeValueRows.length) {
    console.error(
      `Warning: ${INPUT_FREEVALUE_FILE_NAME} has ${freeValueRows.length - freeValueDataIndex} unused data row(s).`,
    );
  }

  return map;
}

function detailRows(
  bRows: string[][],
  assIds: Map<number, string>,
  subjectPkRow: SubjectPkRow,
): string[][] {
  return bRows.slice(1).map((row, index) => {
    const creditAmount = row[COL_CREDIT] ?? "";
    const debitAmount = row[COL_DEBIT] ?? "";
    const period = parsePeriod(row[COL_PERIOD] ?? "");
    const subjectCode = (row[COL_SUBJECT_CODE] ?? "").trim();
    const pkAccsubj = subjectPkRow.pkAccsubjByCode.get(subjectCode);
    if (!pkAccsubj) {
      throw new Error(
        `${INPUT_PK_FILE_NAME} has no SUBJCODE match for ${INPUT_B_FILE_NAME} row ${index + 2}: ${subjectCode}`,
      );
    }
    const pkDetailUuid = String(DETAIL_PK_DETAIL_UUID_START + index).padStart(14, "0");
    const voucherUuid = String(DETAIL_PK_VOUCHER_UUID_START + index);

    return [
      assIds.get(index) ?? "",
      "",
      "",
      "",
      "",
      "",
      "",
      creditAmount,
      "0",
      debitAmount,
      "0",
      row[row.length - 1] ?? "",
      "0",
      "",
      "0",
      "1",
      row[COL_SUMMARY] ?? "",
      "0",
      "0",
      "",
      "",
      "",
      "",
      "",
      creditAmount,
      debitAmount,
      DETAIL_MODIFYFLAG,
      "",
      pkAccsubj,
      subjectPkRow.pkCorp,
      DETAIL_PK_CURRTYPE,
      `1774A9${pkDetailUuid}`,
      DETAIL_PK_GLBOOK,
      subjectPkRow.pkGlorg,
      subjectPkRow.pkGlorgbook,
      "",
      "",
      "",
      "",
      `0001DEFAULT${voucherUuid}`,
      "0",
      "",
      DETAIL_TS,
      isZeroAmount(creditAmount) ? "D" : "C",
      "N",
      "",
      period.month,
      "1",
      period.month,
      "",
      "GL",
      DETAIL_PK_VOUCHERTYPEV,
      formatVoucherDate(row[COL_DATE] ?? ""),
      "",
      "",
      period.year,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "N",
      "",
      "",
      "",
    ];
  });
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

async function processFolder(folderPath: string, utf8Bom: boolean): Promise<number> {
  const bPath = path.join(folderPath, INPUT_B_FILE_NAME);
  const freeValuePath = path.join(folderPath, INPUT_FREEVALUE_FILE_NAME);
  const pkPath = path.join(folderPath, INPUT_PK_FILE_NAME);
  const detailPath = path.join(folderPath, OUTPUT_DETAIL_FILE_NAME);

  const [bRows, freeValueRows, pkRows] = await Promise.all([
    readCsv(bPath),
    readCsv(freeValuePath),
    readCsv(pkPath),
  ]);

  const assIds = assIdsByDataRowIndex(bRows, freeValueRows);
  const subjectPkRow = subjectPkRowFromCsv(pkRows);
  const details = detailRows(bRows, assIds, subjectPkRow);
  await writeLines(
    detailPath,
    [csvRow(DETAIL_HEADER), ...details.map((row) => csvRow(row))],
    utf8Bom,
  );

  console.error("Wrote", details.length + 1, "rows to", detailPath);
  return details.length + 1;
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return 0;
  }

  const folders = await inputFoldersFromArg(args.input);
  if (folders.length === 0) {
    console.error(
      `No folders with ${INPUT_B_FILE_NAME}, ${INPUT_FREEVALUE_FILE_NAME}, and ${INPUT_PK_FILE_NAME} found under ${process.cwd()}`,
    );
    return 1;
  }

  let totalRows = 0;
  for (const folder of folders) {
    console.error("Processing", folder);
    totalRows += await processFolder(folder, args.utf8Bom);
  }

  console.error(
    "Processed",
    folders.length,
    "folder(s),",
    totalRows,
    "detail CSV row(s).",
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
