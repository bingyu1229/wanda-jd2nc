import fs from "node:fs";
import path from "node:path";
import { once } from "node:events";
import { pathToFileURL } from "node:url";

const INPUT_B_FILE_NAME = "b.csv";
const INPUT_DETAIL_FILE_NAME = "GL_DETAIL.csv";
const INPUT_PK_FILE_NAME = "pk.csv";
const INPUT_USER_FILE_NAME = "user.csv";
const OUTPUT_VOUCHER_FILE_NAME = "GL_VOUCHER.csv";

const COL_DATE = 0;
const COL_PREPARED_BY = 11;
const COL_CHECKED_BY = 12;
const COL_POSTED_BY = 13;
const COL_ATTACHMENT = 23;

const VOUCHER_HEADER = [
  "ADDCLASS",
  "ATTACHMENT",
  "CHECKEDDATE",
  "CONTRASTFLAG",
  "CONVERTFLAG",
  "DELETECLASS",
  "DETAILMODFLAG",
  "DISCARDFLAG",
  "DR",
  "ERRMESSAGE",
  "EXPLANATION",
  "FREE1",
  "FREE10",
  "FREE2",
  "FREE3",
  "FREE4",
  "FREE5",
  "FREE6",
  "FREE7",
  "FREE8",
  "FREE9",
  "MODIFYCLASS",
  "MODIFYFLAG",
  "NO",
  "PERIOD",
  "PK_CASHER",
  "PK_CHECKED",
  "PK_CORP",
  "PK_GLBOOK",
  "PK_GLORG",
  "PK_GLORGBOOK",
  "PK_MANAGER",
  "PK_PREPARED",
  "PK_SOB",
  "PK_SOURCEPK",
  "PK_SYSTEM",
  "PK_VOUCHER",
  "PK_VOUCHERTYPE",
  "PREPAREDDATE",
  "SIGNDATE",
  "SIGNFLAG",
  "TALLYDATE",
  "TOTALCREDIT",
  "TOTALDEBIT",
  "TS",
  "VOUCHERKIND",
  "YEAR",
  "ERRMESSAGEH",
  "ISDIFFLAG",
  "OFFERVOUCHER",
];

const VOUCHER_TS = "2026-03-11 9:00:00";
const VOUCHER_TYPE = "0001DEFAULT000000001";

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
  console.error(`Usage: npx tsx script_D.ts [folder-or-b.csv] [options]

Options:
  folder-or-b.csv    Optional folder or b.csv path. When omitted, recursively processes folders with ${INPUT_B_FILE_NAME}, ${INPUT_DETAIL_FILE_NAME}, ${INPUT_PK_FILE_NAME}, and ${INPUT_USER_FILE_NAME}.
  --utf8-bom         Write UTF-8 with BOM for Excel on Windows

Reads ${INPUT_B_FILE_NAME}, ${INPUT_DETAIL_FILE_NAME}, ${INPUT_PK_FILE_NAME}, and ${INPUT_USER_FILE_NAME}; writes ${OUTPUT_VOUCHER_FILE_NAME} next to them.
`);
}

async function findInputFolders(rootDir: string): Promise<string[]> {
  const found: string[] = [];

  async function walk(dir: string): Promise<void> {
    const entries = await fs.promises.readdir(dir, { withFileTypes: true });
    const names = new Set(entries.map((entry) => entry.name));
    if (
      names.has(INPUT_B_FILE_NAME) &&
      names.has(INPUT_DETAIL_FILE_NAME) &&
      names.has(INPUT_PK_FILE_NAME) &&
      names.has(INPUT_USER_FILE_NAME)
    ) {
      found.push(dir);
    }

    for (const entry of entries) {
      if (entry.name === ".git" || entry.name === "node_modules") continue;
      if (entry.isDirectory()) await walk(path.join(dir, entry.name));
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
      .map((value) => (/[",\r\n]/.test(value) ? csvQuoteField(value) : value))
      .join(",") + "\n"
  );
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

function requiredColumnIndex(header: string[], name: string, fileName: string): number {
  const index = header.indexOf(name);
  if (index < 0) throw new Error(`${fileName} is missing ${name}`);
  return index;
}

function userIdsByUserName(userRows: string[][]): Map<string, string> {
  const header = userRows[0] ?? [];
  const userNameCol = requiredColumnIndex(header, "USER_NAME", INPUT_USER_FILE_NAME);
  const cuserIdCol = requiredColumnIndex(header, "CUSERID", INPUT_USER_FILE_NAME);

  const map = new Map<string, string>();
  for (let rowIndex = 1; rowIndex < userRows.length; rowIndex++) {
    const row = userRows[rowIndex];
    const userName = (row[userNameCol] ?? "").trim();
    if (userName) map.set(userName, row[cuserIdCol] ?? "");
  }
  return map;
}

function userIdFor(
  userName: string,
  userIds: Map<string, string>,
  bRowNumber: number,
): string {
  const trimmed = userName.trim();
  if (!trimmed) return "";

  const userId = userIds.get(trimmed);
  if (!userId) {
    throw new Error(
      `${INPUT_USER_FILE_NAME} has no USER_NAME match for ${INPUT_B_FILE_NAME} row ${bRowNumber}: ${trimmed}`,
    );
  }
  return userId;
}

function amountValue(raw: string): number {
  const trimmed = raw.replace(/,/g, "").trim();
  const normalized = trimmed.startsWith("'") ? trimmed.slice(1) : trimmed;
  if (!normalized) return 0;
  const value = Number(normalized);
  if (!Number.isFinite(value)) throw new Error(`Invalid amount: ${raw}`);
  return value;
}

function formatAmount(value: number): string {
  return value.toFixed(2);
}

function formatVoucherDate(date: string): string {
  const match = date.trim().match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!match) return date;
  return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
}

type DetailColumns = {
  creditAmount: number;
  debitAmount: number;
  detailIndex: number;
  periodV: number;
  pkCorp: number;
  pkGlbook: number;
  pkGlorg: number;
  pkGlorgbook: number;
  pkVoucher: number;
  yearV: number;
};

function detailColumns(detailHeader: string[]): DetailColumns {
  return {
    creditAmount: requiredColumnIndex(
      detailHeader,
      "CREDITAMOUNT",
      INPUT_DETAIL_FILE_NAME,
    ),
    debitAmount: requiredColumnIndex(
      detailHeader,
      "DEBITAMOUNT",
      INPUT_DETAIL_FILE_NAME,
    ),
    detailIndex: requiredColumnIndex(
      detailHeader,
      "DETAILINDEX",
      INPUT_DETAIL_FILE_NAME,
    ),
    periodV: requiredColumnIndex(detailHeader, "PERIODV", INPUT_DETAIL_FILE_NAME),
    pkCorp: requiredColumnIndex(detailHeader, "PK_CORP", INPUT_DETAIL_FILE_NAME),
    pkGlbook: requiredColumnIndex(detailHeader, "PK_GLBOOK", INPUT_DETAIL_FILE_NAME),
    pkGlorg: requiredColumnIndex(detailHeader, "PK_GLORG", INPUT_DETAIL_FILE_NAME),
    pkGlorgbook: requiredColumnIndex(
      detailHeader,
      "PK_GLORGBOOK",
      INPUT_DETAIL_FILE_NAME,
    ),
    pkVoucher: requiredColumnIndex(
      detailHeader,
      "PK_VOUCHER",
      INPUT_DETAIL_FILE_NAME,
    ),
    yearV: requiredColumnIndex(detailHeader, "YEARV", INPUT_DETAIL_FILE_NAME),
  };
}

function voucherRows(
  bRows: string[][],
  detailRows: string[][],
  userIds: Map<string, string>,
): string[][] {
  const detailHeader = detailRows[0] ?? [];
  const cols = detailColumns(detailHeader);
  const rows: string[][] = [];
  let currentStart = -1;
  let totalCredit = 0;
  let totalDebit = 0;

  function pushCurrent(): void {
    if (currentStart < 0) return;

    const detailRow = detailRows[currentStart + 1] ?? [];
    const bRow = bRows[currentStart + 1] ?? [];
    const bRowNumber = currentStart + 2;
    const preparedDate = formatVoucherDate(bRow[COL_DATE] ?? "");

    rows.push([
      "",
      bRow[COL_ATTACHMENT] ?? "",
      "",
      "",
      "",
      "",
      "Y",
      "N",
      "0",
      "",
      "",
      detailRow[cols.periodV] ?? "",
      "VOUCHERNEWADD",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "YYY",
      "1",
      detailRow[cols.periodV] ?? "",
      "",
      userIdFor(bRow[COL_CHECKED_BY] ?? "", userIds, bRowNumber),
      detailRow[cols.pkCorp] ?? "",
      detailRow[cols.pkGlbook] ?? "",
      detailRow[cols.pkGlorg] ?? "",
      detailRow[cols.pkGlorgbook] ?? "",
      userIdFor(bRow[COL_POSTED_BY] ?? "", userIds, bRowNumber),
      userIdFor(bRow[COL_PREPARED_BY] ?? "", userIds, bRowNumber),
      "",
      "",
      "GL",
      detailRow[cols.pkVoucher] ?? "",
      VOUCHER_TYPE,
      preparedDate,
      "",
      "Y",
      preparedDate,
      formatAmount(totalCredit),
      formatAmount(totalDebit),
      VOUCHER_TS,
      "0",
      detailRow[cols.yearV] ?? "",
      "",
      "N",
      "",
    ]);
  }

  for (let rowIndex = 1; rowIndex < detailRows.length; rowIndex++) {
    const detailRow = detailRows[rowIndex];
    const dataIndex = rowIndex - 1;
    const detailIndex = (detailRow[cols.detailIndex] ?? "").trim();

    if (detailIndex === "1") {
      pushCurrent();
      currentStart = dataIndex;
      totalCredit = 0;
      totalDebit = 0;
    } else if (currentStart < 0) {
      throw new Error(
        `${INPUT_DETAIL_FILE_NAME} row ${rowIndex + 1} appears before the first DETAILINDEX=1`,
      );
    }

    totalCredit += amountValue(detailRow[cols.creditAmount] ?? "");
    totalDebit += amountValue(detailRow[cols.debitAmount] ?? "");
  }

  pushCurrent();

  return rows;
}

async function processFolder(folderPath: string, utf8Bom: boolean): Promise<number> {
  const bPath = path.join(folderPath, INPUT_B_FILE_NAME);
  const detailPath = path.join(folderPath, INPUT_DETAIL_FILE_NAME);
  const userPath = path.join(folderPath, INPUT_USER_FILE_NAME);
  const outPath = path.join(folderPath, OUTPUT_VOUCHER_FILE_NAME);

  const [bRows, detailRows, userRows] = await Promise.all([
    readCsv(bPath),
    readCsv(detailPath),
    readCsv(userPath),
  ]);

  if (bRows.length !== detailRows.length) {
    throw new Error(
      `${INPUT_B_FILE_NAME} and ${INPUT_DETAIL_FILE_NAME} row counts differ in ${folderPath}: ${bRows.length} vs ${detailRows.length}`,
    );
  }

  const userIds = userIdsByUserName(userRows);
  const vouchers = voucherRows(bRows, detailRows, userIds);
  await writeLines(
    outPath,
    [csvRow(VOUCHER_HEADER), ...vouchers.map((row) => csvRow(row))],
    utf8Bom,
  );

  console.error("Wrote", vouchers.length + 1, "rows to", outPath);
  return vouchers.length + 1;
}

export async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printHelp();
    return 0;
  }

  const folders = await inputFoldersFromArg(args.input);
  if (folders.length === 0) {
    console.error(
      `No folders with ${INPUT_B_FILE_NAME}, ${INPUT_DETAIL_FILE_NAME}, ${INPUT_PK_FILE_NAME}, and ${INPUT_USER_FILE_NAME} found under ${process.cwd()}`,
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
    "voucher CSV row(s).",
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
