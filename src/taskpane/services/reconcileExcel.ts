/*
 * Excel I/O for the reconciliation feature.
 *
 *   - captureCurrentSelection: snapshot the user's current selection into a
 *     persistable descriptor (address, sheet, dimensions, first two rows).
 *   - readRangeValues: fetch the full 2D value grid for a previously-captured range.
 *   - looksLikeHeader / autoDetectColumns: light-touch heuristics for column mapping.
 *   - normalizeRows: turn a 2D grid + ColumnMapping into NormalizedRow[].
 *   - writeReconciliationSheet: produce a new worksheet with a summary block, a
 *     results table, and conditional formatting that color-codes rows by status.
 */

/* global Excel */

import {
  Match,
  MatchingMode,
  NormalizedRow,
  ReconciliationResult,
  Tolerances,
  normalizeDescription,
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
  /** Second row of the selection (used for type-sniffing in autoDetectColumns). */
  sampleRow: unknown[];
}

export interface ColumnMapping {
  dateCol: number;
  amountCol: number;
  descCol: number;
  hasHeader: boolean;
}

/* ----------------------------- Range capture ---------------------------- */

/**
 * Capture the user's current selection. Rejects empty or too-small selections
 * since reconciliation needs at least two rows (header + data, or two data rows).
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
    if (range.rowCount < 2) {
      throw new Error("Select at least two rows (header + data, or two data rows).");
    }
    if (range.columnCount < 3) {
      throw new Error("Select at least three columns so we can map Date, Amount, and Description.");
    }

    descriptor.address = range.address;
    descriptor.sheetName = range.worksheet.name;
    descriptor.a1 = stripSheetPrefix(range.address);
    descriptor.rowCount = range.rowCount;
    descriptor.columnCount = range.columnCount;
    descriptor.firstRow = (range.values[0] as unknown[]) || [];
    descriptor.sampleRow = (range.values[1] as unknown[]) || [];
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

const DATE_HEADER_RE = /date|time|posted|when/i;
const AMOUNT_HEADER_RE = /amount|amt|value|total|debit|credit|sum|price/i;
const DESC_HEADER_RE = /desc|memo|narr|reference|note|payee|merchant|detail/i;

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
 * Suggest column indexes for date, amount, and description.
 * Prefers header-name matches when present; falls back to sample-row type sniffing.
 */
export function autoDetectColumns(
  firstRow: unknown[],
  sampleRow: unknown[],
  hasHeader: boolean
): ColumnMapping {
  const n = firstRow.length;
  let dateCol = -1;
  let amountCol = -1;
  let descCol = -1;

  if (hasHeader) {
    for (let i = 0; i < n; i++) {
      const h = String(firstRow[i] === null || firstRow[i] === undefined ? "" : firstRow[i]).trim();
      if (!h) continue;
      if (dateCol < 0 && DATE_HEADER_RE.test(h)) dateCol = i;
      else if (amountCol < 0 && AMOUNT_HEADER_RE.test(h)) amountCol = i;
      else if (descCol < 0 && DESC_HEADER_RE.test(h)) descCol = i;
    }
  }

  // Type-sniffing fallback using the sample data row.
  const probeRow = hasHeader ? sampleRow : firstRow;
  for (let i = 0; i < n; i++) {
    const v = probeRow[i];
    if (dateCol < 0 && parseDateCell(v) !== null) {
      // Bare numbers can be misread as Excel serial dates; require serial-like magnitude.
      if (typeof v !== "number" || (v > 1000 && v < 2958465)) {
        dateCol = i;
        continue;
      }
    }
    if (amountCol < 0 && i !== dateCol && parseAmountCell(v) !== null && typeof v === "number") {
      amountCol = i;
      continue;
    }
    if (descCol < 0 && i !== dateCol && i !== amountCol && typeof v === "string" && v.trim()) {
      descCol = i;
    }
  }

  // Final fallback: assign leftmost unassigned columns.
  const used: boolean[] = new Array(n);
  if (dateCol >= 0) used[dateCol] = true;
  if (amountCol >= 0) used[amountCol] = true;
  if (descCol >= 0) used[descCol] = true;
  const nextFree = (): number => {
    for (let i = 0; i < n; i++) {
      if (!used[i]) {
        used[i] = true;
        return i;
      }
    }
    return -1;
  };
  if (dateCol < 0) dateCol = nextFree();
  if (amountCol < 0) amountCol = nextFree();
  if (descCol < 0) descCol = nextFree();

  return { dateCol, amountCol, descCol, hasHeader };
}

/* ---------------------- Normalize values into NormalizedRow[] --------------------- */

export function normalizeRows(
  values: unknown[][],
  mapping: ColumnMapping,
  firstExcelRow: number
): NormalizedRow[] {
  const rows: NormalizedRow[] = [];
  const startIdx = mapping.hasHeader ? 1 : 0;
  for (let r = startIdx; r < values.length; r++) {
    const row = values[r] || [];
    const dateVal = row[mapping.dateCol];
    const amountVal = row[mapping.amountCol];
    const descVal = row[mapping.descCol];
    // Skip completely blank rows — common at the bottom of copy-pasted ranges.
    if (isBlank(dateVal) && isBlank(amountVal) && isBlank(descVal)) continue;
    rows.push({
      index: r - startIdx,
      excelRow: firstExcelRow + r,
      dateMs: parseDateCell(dateVal),
      amount: parseAmountCell(amountVal),
      descriptionNorm: normalizeDescription(descVal),
      dateDisplay: formatDateCell(dateVal),
      amountDisplay: formatAmountCell(amountVal),
      descriptionDisplay: descVal === null || descVal === undefined ? "" : String(descVal),
    });
  }
  return rows;
}

function isBlank(v: unknown): boolean {
  return v === null || v === undefined || v === "";
}

function formatDateCell(v: unknown): string {
  const ms = parseDateCell(v);
  if (ms === null) return v === null || v === undefined ? "" : String(v);
  return new Date(ms).toISOString().slice(0, 10);
}

function formatAmountCell(v: unknown): number | string {
  const n = parseAmountCell(v);
  if (n === null) return v === null || v === undefined ? "" : String(v);
  return n;
}

/* -------------------------- Results worksheet output ------------------------- */

const COLOR_MATCHED = "#DFF6DD";
const COLOR_POSSIBLE = "#FFF4CE";
const COLOR_UNMATCHED = "#FDE7E9";
const COLOR_HEADER_FILL = "#F3F2F1";
const COLOR_TITLE_FILL = "#0F6CBD";

export interface WriteResultsOptions {
  mode: MatchingMode;
  tolerances: Tolerances;
  descA: RangeDescriptor;
  descB: RangeDescriptor;
}

/**
 * Write the reconciliation result to a new worksheet with a summary block, a
 * results table, and three conditional-format rules coloring rows by status.
 * Returns the created sheet's final name.
 */
export async function writeReconciliationSheet(
  result: ReconciliationResult,
  opts: WriteResultsOptions
): Promise<string> {
  const baseName = sanitizeSheetName(`Reconciliation ${new Date().toISOString().slice(0, 10)}`);
  let finalName = baseName;

  const headerCols = [
    "Side",
    "Row Ref",
    "Date",
    "Amount",
    "Description",
    "Status",
    "Confidence %",
    "Difference",
    "Notes",
  ];
  // Status column in the results table is the 6th (F). Used by conditional-format formulas.
  const STATUS_COL_LETTER = "F";
  const AMOUNT_COL_INDEX = 3;

  // Build output rows: matches (sorted by confidence desc), then unmatched A, then unmatched B.
  const outRows: (string | number)[][] = [];
  const sortedMatches = result.matches.slice().sort((a, b) => b.confidence - a.confidence);
  for (let k = 0; k < sortedMatches.length; k++) {
    const m = sortedMatches[k];
    outRows.push(buildMatchRow("A", m.a, m, `Paired with B row ${m.b.excelRow}`));
    outRows.push(buildMatchRow("B", m.b, m, `Paired with A row ${m.a.excelRow}`));
  }
  for (let i = 0; i < result.unmatchedA.length; i++) {
    outRows.push(buildUnmatchedRow("A", result.unmatchedA[i]));
  }
  for (let j = 0; j < result.unmatchedB.length; j++) {
    outRows.push(buildUnmatchedRow("B", result.unmatchedB[j]));
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
    const summaryPairs: Array<[string, string]> = [
      ["Reconciliation Results", `Generated ${new Date().toLocaleString()}`],
      ["Matching Mode", capitalize(opts.mode)],
      [
        "Tolerances",
        `Amount: $${opts.tolerances.amountFixed.toFixed(2)} or ${opts.tolerances.amountPct}%  ·  Date: ±${opts.tolerances.dateDays} day(s)`,
      ],
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

      // Format the Amount column as currency.
      const amountRange = sheet.getRangeByIndexes(dataStart, AMOUNT_COL_INDEX, outRows.length, 1);
      const amountFmt: string[][] = new Array(outRows.length);
      for (let i = 0; i < outRows.length; i++) amountFmt[i] = ["$#,##0.00;[Red]($#,##0.00)"];
      amountRange.numberFormat = amountFmt;

      // Conditional formatting: color whole rows based on the Status column.
      const rules: Array<{ status: string; color: string }> = [
        { status: "Matched", color: COLOR_MATCHED },
        { status: "Possible Match", color: COLOR_POSSIBLE },
        { status: "Unmatched", color: COLOR_UNMATCHED },
      ];
      for (let i = 0; i < rules.length; i++) {
        const rule = rules[i];
        const cf = dataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = `=$${STATUS_COL_LETTER}${dataStart + 1}="${rule.status}"`;
        cf.custom.format.fill.color = rule.color;
      }
    }

    /* Autofit and freeze summary + header. */
    const totalRows = Math.max(tableStartRow + 1 + outRows.length, 1);
    sheet.getRangeByIndexes(0, 0, totalRows, headerCols.length).format.autofitColumns();
    sheet.freezePanes.freezeRows(tableStartRow + 1);

    sheet.activate();
    await context.sync();
  });

  return finalName;
}

function buildMatchRow(
  side: "A" | "B",
  row: NormalizedRow,
  m: Match,
  pairingNote: string
): (string | number)[] {
  // The amount difference is reported from the current side's perspective.
  const diffAmount = side === "A" ? m.amountDiff : -m.amountDiff;
  const diffDays = side === "A" ? m.dateDiffDays : -m.dateDiffDays;
  return [
    side,
    row.excelRow,
    row.dateDisplay,
    row.amountDisplay,
    row.descriptionDisplay,
    m.status,
    Math.round(m.confidence * 100),
    describeDifference(diffAmount, diffDays, m.descSimilarity),
    `${pairingNote} · ${m.notes}`,
  ];
}

function buildUnmatchedRow(side: "A" | "B", row: NormalizedRow): (string | number)[] {
  return [
    side,
    row.excelRow,
    row.dateDisplay,
    row.amountDisplay,
    row.descriptionDisplay,
    "Unmatched",
    0,
    "—",
    "No candidate above the possible-match threshold",
  ];
}

function describeDifference(amountDiff: number, dateDiffDays: number, descSim: number): string {
  const parts: string[] = [];
  if (Math.abs(amountDiff) > 1e-9) {
    const sign = amountDiff >= 0 ? "+" : "-";
    parts.push(`${sign}$${Math.abs(amountDiff).toFixed(2)}`);
  } else {
    parts.push("$0.00");
  }
  if (dateDiffDays !== 0) {
    const sign = dateDiffDays > 0 ? "+" : "";
    parts.push(`${sign}${dateDiffDays}d`);
  }
  if (descSim < 1) {
    parts.push(`desc ${Math.round(descSim * 100)}%`);
  }
  return parts.join(" / ");
}

function sanitizeSheetName(name: string): string {
  // Excel: max 31 chars; cannot contain \ / ? * [ ] :
  const cleaned = name.replace(/[\\/?*[\]:]/g, "_").trim();
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned || "Sheet";
}

function capitalize(s: string): string {
  return s.length ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}
