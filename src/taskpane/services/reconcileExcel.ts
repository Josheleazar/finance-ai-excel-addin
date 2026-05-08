/*
 * Excel I/O for the reconciliation feature.
 *
 *   - captureCurrentSelection: snapshot the user's current selection into a
 *     persistable descriptor (address, sheet, dimensions, first two rows).
 *   - readRangeValues: fetch the full 2D value grid for a captured range.
 *   - looksLikeHeader / detectColumnForField: heuristics for column auto-mapping.
 *   - buildRawRows: trim the optional header row, skip blank rows, preserve full
 *     cell arrays so the matching engine can read any column by index.
 *   - writeReconciliationSheet: produce a new worksheet whose results table has
 *     one column per user-configured field plus Status / Confidence / Notes,
 *     with per-status conditional formatting on the row.
 */

/* global Excel */

import {
  FieldConfig,
  Match,
  RawRow,
  ReconciliationResult,
  Sensitivity,
  parseAmountCell,
  parseDateCell,
} from "./reconciliation";

export interface RangeDescriptor {
  /** Full address including sheet name, e.g. "Sheet1!A1:D25". */
  address: string;
  /** Worksheet name portion of the address. */
  sheetName: string;
  /** A1 notation without the sheet prefix, e.g. "A1:D25". */
  a1: string;
  rowCount: number;
  columnCount: number;
  /** First row of the selection (used for header detection & column dropdowns). */
  firstRow: unknown[];
  /** Second row of the selection (used for type-sniffing in auto-detection). */
  sampleRow: unknown[];
}

/* ----------------------------- Range capture ---------------------------- */

/**
 * Capture the user's current selection. Accepts any non-empty range — per-field
 * column mapping happens after capture, so narrow single-column ranges are fine.
 */
export async function captureCurrentSelection(): Promise<RangeDescriptor> {
  const descriptor: RangeDescriptor = {
    address: "",
    sheetName: "",
    a1: "",
    rowCount: 0,
    columnCount: 0,
    firstRow: [],
    sampleRow: [],
  };

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, rowCount, columnCount, values, worksheet/name");
    await context.sync();

    if (!range.rowCount || !range.columnCount) {
      throw new Error("No range is selected. Highlight your data and try again.");
    }

    descriptor.address = range.address;
    descriptor.sheetName = range.worksheet.name;
    descriptor.a1 = stripSheetPrefix(range.address);
    descriptor.rowCount = range.rowCount;
    descriptor.columnCount = range.columnCount;
    descriptor.firstRow = (range.values[0] as unknown[]) || [];
    descriptor.sampleRow = range.rowCount > 1 ? (range.values[1] as unknown[]) || [] : [];
  });

  return descriptor;
}

/** Read values for a previously-captured range. Returns the full 2D value grid. */
export async function readRangeValues(desc: RangeDescriptor): Promise<unknown[][]> {
  let values: unknown[][] = [];
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(desc.sheetName);
    sheet.load("name");
    await context.sync();
    if (sheet.isNullObject) {
      throw new Error(`Worksheet "${desc.sheetName}" no longer exists.`);
    }
    const range = sheet.getRange(desc.a1);
    range.load("values");
    await context.sync();
    values = range.values as unknown[][];
  });
  return values;
}

function stripSheetPrefix(fullAddress: string): string {
  // Excel returns e.g. "Sheet1!A1:D25" or "'My Sheet'!A1:D25".
  const bang = fullAddress.lastIndexOf("!");
  return bang < 0 ? fullAddress : fullAddress.slice(bang + 1);
}

/* ------------------- Header detection & column auto-mapping ------------------- */

const DATE_HEADER_RE = /date|time|posted|period|when/i;
const NUMERIC_HEADER_RE =
  /amount|amt|value|total|debit|credit|sum|price|qty|quantity|units|shares|balance|cost/i;
const TEXT_HEADER_RE =
  /desc|memo|narr|reference|ref|note|payee|merchant|detail|vendor|customer|name|invoice|id|sku|ticker/i;

/** A row "looks like a header" if every cell is either blank or a non-numeric string. */
export function looksLikeHeader(firstRow: unknown[]): boolean {
  if (!firstRow.length) return false;
  let stringCount = 0;
  for (let i = 0; i < firstRow.length; i++) {
    const v = firstRow[i];
    if (typeof v === "number") return false;
    if (v instanceof Date) return false;
    if (typeof v === "string" && v.trim()) stringCount++;
  }
  return stringCount >= 1;
}

/**
 * Suggest a column index for a field based on its type/label and the captured
 * range. Tries the field label, then type-specific header regex, then sample-
 * row type-sniffing. Returns -1 when nothing matches so callers can leave the
 * field unset rather than silently guess.
 *
 * Pass `skip` to avoid returning column indexes already claimed by other fields.
 */
export function detectColumnForField(
  field: { type: string; label: string },
  desc: RangeDescriptor,
  hasHeader: boolean,
  skip?: { [col: number]: boolean }
): number {
  const firstRow = desc.firstRow;
  const n = firstRow.length;
  if (n === 0) return -1;
  const taken = skip || {};

  // 1. Direct label match in the header row (user-customized labels win).
  if (hasHeader && field.label) {
    const want = field.label.trim().toLowerCase();
    if (want) {
      for (let i = 0; i < n; i++) {
        if (taken[i]) continue;
        const h = String(firstRow[i] == null ? "" : firstRow[i])
          .trim()
          .toLowerCase();
        if (h && (h === want || h.indexOf(want) >= 0 || want.indexOf(h) >= 0)) return i;
      }
    }
  }

  // 2. Type-specific header regex.
  const regex =
    field.type === "date"
      ? DATE_HEADER_RE
      : field.type === "numeric"
        ? NUMERIC_HEADER_RE
        : TEXT_HEADER_RE;
  if (hasHeader) {
    for (let i = 0; i < n; i++) {
      if (taken[i]) continue;
      const h = String(firstRow[i] == null ? "" : firstRow[i]).trim();
      if (h && regex.test(h)) return i;
    }
  }

  // 3. Sniff sample row types.
  const probe = hasHeader ? desc.sampleRow : firstRow;
  for (let i = 0; i < probe.length; i++) {
    if (taken[i]) continue;
    const v = probe[i];
    if (field.type === "date") {
      if (parseDateCell(v) !== null) {
        if (typeof v !== "number" || (v > 1000 && v < 2958465)) return i;
      }
    } else if (field.type === "numeric") {
      if (typeof v === "number" && isFinite(v)) return i;
    } else {
      if (typeof v === "string" && v.trim()) return i;
    }
  }

  // Nothing matched — leave unset so the user knows to pick a column.
  return -1;
}

/* ---------------------- Row normalization ---------------------- */

/**
 * Turn a 2D value grid into RawRow[] — trims an optional header row and skips
 * fully-blank rows (common at the bottom of copy-pasted ranges). Cells are
 * preserved as-is so fields can index into them later.
 */
export function buildRawRows(
  values: unknown[][],
  hasHeader: boolean,
  firstExcelRow: number
): RawRow[] {
  const rows: RawRow[] = [];
  const startIdx = hasHeader ? 1 : 0;
  let dataIdx = 0;
  for (let r = startIdx; r < values.length; r++) {
    const row = values[r] || [];
    if (isRowBlank(row)) continue;
    rows.push({
      excelRow: firstExcelRow + r,
      index: dataIdx++,
      cells: row,
    });
  }
  return rows;
}

function isRowBlank(row: unknown[]): boolean {
  for (let i = 0; i < row.length; i++) {
    const v = row[i];
    if (v !== null && v !== undefined && v !== "") return false;
  }
  return true;
}

/* -------------------------- Cell display helpers -------------------------- */

function displayCell(v: unknown, fieldType: string): string | number {
  if (v === null || v === undefined || v === "") return "";
  if (fieldType === "date") {
    const ms = parseDateCell(v);
    if (ms !== null) return new Date(ms).toISOString().slice(0, 10);
    return String(v);
  }
  if (fieldType === "numeric") {
    const n = parseAmountCell(v);
    if (n !== null) return n;
    return String(v);
  }
  if (typeof v === "number" || typeof v === "string") return v;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
}

/* -------------------------- Results worksheet output ------------------------- */

const COLOR_MATCHED = "#DFF6DD";
const COLOR_POSSIBLE = "#FFF4CE";
const COLOR_UNMATCHED = "#FDE7E9";
const COLOR_HEADER_FILL = "#F3F2F1";
const COLOR_TITLE_FILL = "#0F6CBD";

export interface WriteResultsOptions {
  sensitivity: Sensitivity;
  fields: FieldConfig[];
  descA: RangeDescriptor;
  descB: RangeDescriptor;
}

/**
 * Write the reconciliation result to a new worksheet. Columns are:
 *   Side | Row Ref | <one per field> | Status | Confidence % | Notes
 * with conditional-format rules coloring rows green / yellow / red by Status.
 */
export async function writeReconciliationSheet(
  result: ReconciliationResult,
  opts: WriteResultsOptions
): Promise<string> {
  const baseName = sanitizeSheetName(`Reconciliation ${new Date().toISOString().slice(0, 10)}`);
  let finalName = baseName;

  const fields = opts.fields;

  // Dynamic columns: two fixed header cells (Side, Row Ref) + one per field + three trailing.
  const fieldHeaders = fields.map((f) => f.label || `Field ${f.id}`);
  const headerCols: string[] = ["Side", "Row Ref"]
    .concat(fieldHeaders)
    .concat(["Status", "Confidence %", "Notes"]);
  const statusColIndex = 2 + fields.length;
  const statusColLetter = columnIndexToLetter(statusColIndex);
  const numericCols = collectNumericColIndexes(fields);

  // Build output rows: matches (paired, confidence desc), then unmatched A, then unmatched B.
  const outRows: (string | number)[][] = [];
  const sortedMatches = result.matches.slice().sort((a, b) => b.confidence - a.confidence);
  for (let k = 0; k < sortedMatches.length; k++) {
    const m = sortedMatches[k];
    outRows.push(buildPairedRow("A", m.a, m, m.b.excelRow, fields));
    outRows.push(buildPairedRow("B", m.b, m, m.a.excelRow, fields));
  }
  for (let i = 0; i < result.unmatchedA.length; i++) {
    outRows.push(buildUnmatchedRow("A", result.unmatchedA[i], fields));
  }
  for (let j = 0; j < result.unmatchedB.length; j++) {
    outRows.push(buildUnmatchedRow("B", result.unmatchedB[j], fields));
  }

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const existing: { [name: string]: boolean } = {};
    for (let i = 0; i < sheets.items.length; i++) existing[sheets.items[i].name] = true;
    if (existing[finalName]) {
      let i = 2;
      while (existing[`${baseName} (${i})`]) i++;
      finalName = `${baseName} (${i})`;
    }
    const sheet = sheets.add(finalName);

    /* Summary block. */
    const fieldSummary =
      fields.length === 0
        ? "— no fields configured —"
        : fields
            .map((f) => {
              const reqTag = f.required ? (f.downgradeOnFail ? " [req↓]" : " [req!]") : "";
              return `${f.label} (${f.type}, ${f.weight}${reqTag})`;
            })
            .join(", ");

    const summaryPairs: Array<[string, string]> = [
      ["Reconciliation Results", `Generated ${new Date().toLocaleString()}`],
      ["Sensitivity", capitalize(opts.sensitivity)],
      ["Fields", fieldSummary],
      [
        "Source A",
        `${opts.descA.address}  (${opts.descA.rowCount} rows × ${opts.descA.columnCount} cols)`,
      ],
      [
        "Source B",
        `${opts.descB.address}  (${opts.descB.rowCount} rows × ${opts.descB.columnCount} cols)`,
      ],
      [
        "Summary",
        `Matched: ${result.summary.matched}  ·  Possible: ${result.summary.possible}  ·  Unmatched A: ${result.summary.unmatchedA}  ·  Unmatched B: ${result.summary.unmatchedB}`,
      ],
    ];
    for (let r = 0; r < summaryPairs.length; r++) {
      const row = sheet.getRangeByIndexes(r, 0, 1, 2);
      row.values = [[summaryPairs[r][0], summaryPairs[r][1]]];
      const labelCell = sheet.getRangeByIndexes(r, 0, 1, 1);
      labelCell.format.font.bold = true;
      if (r === 0) {
        const titleRow = sheet.getRangeByIndexes(r, 0, 1, 2);
        titleRow.format.fill.color = COLOR_TITLE_FILL;
        titleRow.format.font.color = "#ffffff";
        titleRow.format.font.bold = true;
      }
    }

    /* Results table. */
    const tableStartRow = summaryPairs.length + 1; // blank spacer row after summary
    const headerRange = sheet.getRangeByIndexes(tableStartRow, 0, 1, headerCols.length);
    headerRange.values = [headerCols];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = COLOR_HEADER_FILL;

    if (outRows.length > 0) {
      const dataStart = tableStartRow + 1;
      const dataRange = sheet.getRangeByIndexes(dataStart, 0, outRows.length, headerCols.length);
      dataRange.values = outRows;

      // Currency-format each numeric field column.
      for (let nc = 0; nc < numericCols.length; nc++) {
        const ci = numericCols[nc];
        const colRange = sheet.getRangeByIndexes(dataStart, ci, outRows.length, 1);
        const fmt: string[][] = new Array(outRows.length);
        for (let i = 0; i < outRows.length; i++) fmt[i] = ["#,##0.00;[Red](#,##0.00)"];
        colRange.numberFormat = fmt;
      }

      // Conditional formatting: color whole rows by Status column.
      const rules: Array<{ status: string; color: string }> = [
        { status: "Matched", color: COLOR_MATCHED },
        { status: "Possible Match", color: COLOR_POSSIBLE },
        { status: "Unmatched", color: COLOR_UNMATCHED },
      ];
      for (let i = 0; i < rules.length; i++) {
        const rule = rules[i];
        const cf = dataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = `=$${statusColLetter}${dataStart + 1}="${rule.status}"`;
        cf.custom.format.fill.color = rule.color;
      }
    }

    const totalRows = Math.max(tableStartRow + 1 + outRows.length, 1);
    sheet.getRangeByIndexes(0, 0, totalRows, headerCols.length).format.autofitColumns();
    sheet.freezePanes.freezeRows(tableStartRow + 1);

    sheet.activate();
    await context.sync();
  });

  return finalName;
}

function buildPairedRow(
  side: "A" | "B",
  row: RawRow,
  m: Match,
  pairedToExcelRow: number,
  fields: FieldConfig[]
): (string | number)[] {
  const vals: (string | number)[] = [side, row.excelRow];
  for (let k = 0; k < fields.length; k++) {
    const f = fields[k];
    const col = side === "A" ? f.colA : f.colB;
    vals.push(col >= 0 && col < row.cells.length ? displayCell(row.cells[col], f.type) : "");
  }
  vals.push(m.status);
  vals.push(Math.round(m.confidence * 100));
  vals.push(`Paired with ${side === "A" ? "B" : "A"} row ${pairedToExcelRow} · ${m.notes}`);
  return vals;
}

function buildUnmatchedRow(
  side: "A" | "B",
  row: RawRow,
  fields: FieldConfig[]
): (string | number)[] {
  const vals: (string | number)[] = [side, row.excelRow];
  for (let k = 0; k < fields.length; k++) {
    const f = fields[k];
    const col = side === "A" ? f.colA : f.colB;
    vals.push(col >= 0 && col < row.cells.length ? displayCell(row.cells[col], f.type) : "");
  }
  vals.push("Unmatched");
  vals.push(0);
  vals.push("No candidate above the possible-match threshold");
  return vals;
}

function collectNumericColIndexes(fields: FieldConfig[]): number[] {
  const result: number[] = [];
  for (let k = 0; k < fields.length; k++) {
    if (fields[k].type === "numeric") result.push(2 + k);
  }
  return result;
}

function sanitizeSheetName(name: string): string {
  // Excel: max 31 chars; cannot contain \ / ? * [ ] :
  const cleaned = name.replace(/[\\/?*[\]:]/g, "_").trim();
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned || "Sheet";
}

function capitalize(s: string): string {
  return s.length ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

export function columnIndexToLetter(i: number): string {
  let s = "";
  let n = i;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}
