/*
 * Excel I/O for the Portfolio Tracker feature.
 *
 * Single source of truth: a worksheet named PORTFOLIO_SHEET_NAME holds one
 * row per holding, with a fixed column order. The task pane reads this sheet
 * on mount and writes the entire sheet back after every edit. This keeps the
 * UI and sheet in lock-step without needing events or change-tracking.
 *
 * We also own the "reconcile portfolio vs broker selection" flow here because
 * it is very specific to the portfolio schema (Symbol + Quantity + Cost Basis
 * comparison with a purpose-built output table).
 */

/* global Excel */

import {
  Holding,
  computeHoldingMetrics,
  computeTotals,
  normalizeSymbol,
  parseLooseNumber,
} from "./portfolio";
import { columnIndexToLetter } from "./reconcileExcel";

export const PORTFOLIO_SHEET_NAME = "Portfolio";

/** Column order on the Portfolio worksheet. Used for both read and write so
 *  the two paths stay in sync. */
export const PORTFOLIO_COLUMNS = [
  "Symbol",
  "Company Name",
  "Sector",
  "Quantity",
  "Avg Cost Basis",
  "Current Price",
  "Market Value",
  "Unrealized P&L ($)",
  "Unrealized P&L (%)",
  "Weight %",
  "Purchase Date",
  "Account Type",
  "Last Refreshed",
] as const;

const COLOR_HEADER_FILL = "#F3F2F1";
const COLOR_TOTAL_FILL = "#EAF3FB";
const COLOR_POSITIVE = "#107C10";
const COLOR_NEGATIVE = "#A4262C";

/* --------------------------- Sheet read --------------------------- */

/**
 * Load holdings from the Portfolio worksheet. Returns an empty array if the
 * sheet doesn't exist yet (first-run). Rows missing a Symbol are skipped.
 */
export async function readPortfolioFromSheet(): Promise<Holding[]> {
  const result: Holding[] = [];
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.getItemOrNullObject(PORTFOLIO_SHEET_NAME);
    sheet.load("name");
    await context.sync();
    if (sheet.isNullObject) return;

    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("values, rowCount, columnCount");
    await context.sync();
    if (used.isNullObject || !used.rowCount || used.rowCount < 2) return;

    const values = used.values as unknown[][];
    // Row 0 is the header. Stop reading when we hit the "Totals" marker.
    for (let r = 1; r < values.length; r++) {
      const row = values[r] || [];
      const firstCell = row[0];
      if (
        typeof firstCell === "string" &&
        firstCell.trim().toLowerCase() === "totals"
      ) {
        break;
      }
      const symbol = normalizeSymbol(firstCell);
      if (!symbol) continue;
      const h: Holding = {
        symbol,
        companyName: cellToOptionalString(row[1]),
        sector: cellToOptionalString(row[2]),
        quantity: parseLooseNumber(row[3]) || 0,
        avgCostBasis: parseLooseNumber(row[4]) || 0,
      };
      const price = parseLooseNumber(row[5]);
      if (price !== null) h.currentPrice = price;
      const purchaseDate = cellToOptionalString(row[10]);
      if (purchaseDate) h.purchaseDate = purchaseDate;
      const accountType = cellToOptionalString(row[11]);
      if (accountType) h.accountType = accountType;
      const refreshed = row[12];
      if (typeof refreshed === "number" && refreshed > 0) {
        h.lastRefreshedAt = refreshed;
      } else if (typeof refreshed === "string" && refreshed.trim()) {
        const t = Date.parse(refreshed);
        if (!isNaN(t)) h.lastRefreshedAt = t;
      }
      result.push(h);
    }
  });
  return result;
}

function cellToOptionalString(v: unknown): string | undefined {
  if (v === null || v === undefined) return undefined;
  const s = String(v).trim();
  return s ? s : undefined;
}

/* --------------------------- Sheet write --------------------------- */

/**
 * Write the current holdings list to the Portfolio worksheet, overwriting
 * any previous content. A totals row is appended at the bottom.
 *
 * The write happens inside a single Excel.run / context.sync for speed.
 * The sheet is created on first call and activated so the user can see it.
 */
export async function writePortfolioToSheet(holdings: Holding[]): Promise<void> {
  const totals = computeTotals(holdings);

  const dataRows: (string | number)[][] = new Array(holdings.length);
  for (let i = 0; i < holdings.length; i++) {
    const h = holdings[i];
    const m = computeHoldingMetrics(h, totals.marketValue);
    dataRows[i] = [
      h.symbol,
      h.companyName || "",
      h.sector || "",
      h.quantity,
      h.avgCostBasis,
      h.currentPrice === undefined ? "" : h.currentPrice,
      m.marketValue,
      m.unrealizedPnl,
      m.unrealizedPnlPct,
      m.weightPct,
      h.purchaseDate || "",
      h.accountType || "",
      h.lastRefreshedAt ? new Date(h.lastRefreshedAt).toISOString() : "",
    ];
  }

  const headerRow: string[] = [...PORTFOLIO_COLUMNS];
  const totalsRow: (string | number)[] = [
    "Totals",
    "",
    "",
    "", // quantity aggregation would be misleading across symbols
    "",
    "",
    totals.marketValue,
    totals.unrealizedPnl,
    totals.unrealizedPnlPct,
    100,
    "",
    "",
    "",
  ];

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItemOrNullObject(PORTFOLIO_SHEET_NAME);
    sheet.load("name");
    await context.sync();

    if (sheet.isNullObject) {
      sheet = sheets.add(PORTFOLIO_SHEET_NAME);
    } else {
      // Clear prior content before re-writing so deletions propagate.
      // `getUsedRangeOrNullObject` handles the edge case of an empty sheet.
      const prior = sheet.getUsedRangeOrNullObject(true);
      prior.load("isNullObject");
      await context.sync();
      if (!prior.isNullObject) prior.clear();
    }

    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headerRow.length);
    headerRange.values = [headerRow];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = COLOR_HEADER_FILL;

    if (dataRows.length > 0) {
      const dataRange = sheet.getRangeByIndexes(1, 0, dataRows.length, headerRow.length);
      dataRange.values = dataRows;
      applyNumberFormats(sheet, 1, dataRows.length, headerRow.length);
      applyPnlConditionalFormat(sheet, 1, dataRows.length);
    }

    const totalsRowIdx = 1 + dataRows.length;
    const totalsRange = sheet.getRangeByIndexes(totalsRowIdx, 0, 1, headerRow.length);
    totalsRange.values = [totalsRow];
    totalsRange.format.font.bold = true;
    totalsRange.format.fill.color = COLOR_TOTAL_FILL;
    applyNumberFormats(sheet, totalsRowIdx, 1, headerRow.length);

    const fullRange = sheet.getRangeByIndexes(0, 0, totalsRowIdx + 1, headerRow.length);
    fullRange.format.autofitColumns();
    sheet.freezePanes.freezeRows(1);
    sheet.activate();
    await context.sync();
  });
}

function applyNumberFormats(
  sheet: Excel.Worksheet,
  startRow: number,
  rowCount: number,
  colCount: number
): void {
  // Column indices from PORTFOLIO_COLUMNS:
  //   3 Quantity, 4 Cost Basis, 5 Current Price, 6 Market Value,
  //   7 P&L $, 8 P&L %, 9 Weight %
  const fmtMoney = "$#,##0.00;[Red]($#,##0.00)";
  const fmtMoneyNoSign = "#,##0.0000";
  const fmtWeight = "0.00\"%\"";

  const colFormats: Array<{ col: number; format: string }> = [
    { col: 3, format: "#,##0.####" },
    { col: 4, format: fmtMoneyNoSign },
    { col: 5, format: "$#,##0.00" },
    { col: 6, format: fmtMoney },
    { col: 7, format: fmtMoney },
    { col: 8, format: "0.00\"%\";[Red]-0.00\"%\"" },
    { col: 9, format: fmtWeight },
  ];

  for (let i = 0; i < colFormats.length; i++) {
    const { col, format } = colFormats[i];
    if (col >= colCount) continue;
    const range = sheet.getRangeByIndexes(startRow, col, rowCount, 1);
    const grid: string[][] = new Array(rowCount);
    for (let r = 0; r < rowCount; r++) grid[r] = [format];
    range.numberFormat = grid;
  }
}

/**
 * Color P&L ($) and P&L (%) cells green / red based on sign. Uses a custom
 * conditional format anchored to the P&L $ column (col index 7).
 */
function applyPnlConditionalFormat(
  sheet: Excel.Worksheet,
  startRow: number,
  rowCount: number
): void {
  if (rowCount <= 0) return;
  // P&L ($) and P&L (%) columns
  for (let c = 7; c <= 8; c++) {
    const range = sheet.getRangeByIndexes(startRow, c, rowCount, 1);
    const colLetter = columnIndexToLetter(c);
    const posRule = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    posRule.custom.rule.formula = `=${colLetter}${startRow + 1}>0`;
    posRule.custom.format.font.color = COLOR_POSITIVE;
    const negRule = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    negRule.custom.rule.formula = `=${colLetter}${startRow + 1}<0`;
    negRule.custom.format.font.color = COLOR_NEGATIVE;
  }
}

/* --------------------- Broker range capture --------------------- */

export interface BrokerRange {
  address: string;
  sheetName: string;
  a1: string;
  rowCount: number;
  columnCount: number;
  headers: string[];
  rows: unknown[][];
}

/**
 * Snapshot the user's current Excel selection and eagerly read all values.
 * First row is assumed to be a header (broker exports always have one).
 */
export async function captureBrokerSelection(): Promise<BrokerRange> {
  const out: BrokerRange = {
    address: "",
    sheetName: "",
    a1: "",
    rowCount: 0,
    columnCount: 0,
    headers: [],
    rows: [],
  };
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, rowCount, columnCount, values, worksheet/name");
    await context.sync();
    if (!range.rowCount || !range.columnCount) {
      throw new Error("No range is selected. Highlight your broker data and retry.");
    }
    if (range.rowCount < 2) {
      throw new Error(
        "Selection only has one row. Include a header row and at least one data row."
      );
    }
    out.address = range.address;
    out.sheetName = range.worksheet.name;
    const bang = range.address.lastIndexOf("!");
    out.a1 = bang < 0 ? range.address : range.address.slice(bang + 1);
    out.rowCount = range.rowCount;
    out.columnCount = range.columnCount;
    const values = range.values as unknown[][];
    out.headers = (values[0] || []).map((v) =>
      v === null || v === undefined ? "" : String(v).trim()
    );
    out.rows = values.slice(1);
  });
  return out;
}

/* -------------------- Portfolio reconciliation -------------------- */

/** Per-row outcome when reconciling the portfolio against a broker range. */
export type PortfolioMatchStatus = "Matched" | "Warning" | "Error";

export interface PortfolioReconRow {
  symbol: string;
  portfolioQty: number | "";
  brokerQty: number | "";
  qtyDiff: number | "";
  portfolioCost: number | "";
  brokerCost: number | "";
  costDiff: number | "";
  status: PortfolioMatchStatus;
  note: string;
}

export interface PortfolioReconOptions {
  qtyTolerance: number; // absolute, e.g. 0.0001 shares
  costTolerance: number; // absolute $/share, e.g. 0.01
  /** Mapping of broker columns: index of symbol, quantity, cost basis (or -1). */
  symbolCol: number;
  quantityCol: number;
  costCol: number;
}

/**
 * Compare the portfolio (our side) against a broker-supplied range. Matches
 * by symbol. Produces one row per symbol that appears on either side:
 *
 *   - Both sides, qty & cost within tolerance                → Matched
 *   - Both sides but qty or cost differs beyond tolerance    → Warning
 *   - Symbol missing from one side                           → Error
 */
export function reconcilePortfolio(
  holdings: Holding[],
  broker: BrokerRange,
  opts: PortfolioReconOptions
): PortfolioReconRow[] {
  // Aggregate broker rows by symbol so duplicate lots collapse into one entry.
  const brokerMap: { [sym: string]: { qty: number; costWeightSum: number } } = {};
  for (let i = 0; i < broker.rows.length; i++) {
    const row = broker.rows[i] || [];
    const sym = normalizeSymbol(opts.symbolCol >= 0 ? row[opts.symbolCol] : "");
    if (!sym) continue;
    const qty = parseLooseNumber(opts.quantityCol >= 0 ? row[opts.quantityCol] : 0) || 0;
    const cost = parseLooseNumber(opts.costCol >= 0 ? row[opts.costCol] : 0) || 0;
    if (!brokerMap[sym]) brokerMap[sym] = { qty: 0, costWeightSum: 0 };
    brokerMap[sym].qty += qty;
    // Weighted-average cost basis when multiple rows exist for the same symbol.
    brokerMap[sym].costWeightSum += qty * cost;
  }
  // Resolve weighted-average cost.
  const brokerBySym: { [sym: string]: { qty: number; cost: number } } = {};
  for (const sym in brokerMap) {
    const b = brokerMap[sym];
    brokerBySym[sym] = {
      qty: b.qty,
      cost: b.qty !== 0 ? b.costWeightSum / b.qty : 0,
    };
  }

  const portfolioBySym: { [sym: string]: Holding } = {};
  for (let i = 0; i < holdings.length; i++) {
    portfolioBySym[holdings[i].symbol] = holdings[i];
  }

  const allSymbols: string[] = [];
  const seen: { [s: string]: boolean } = {};
  for (let i = 0; i < holdings.length; i++) {
    const s = holdings[i].symbol;
    if (!seen[s]) {
      allSymbols.push(s);
      seen[s] = true;
    }
  }
  for (const s in brokerBySym) {
    if (!seen[s]) {
      allSymbols.push(s);
      seen[s] = true;
    }
  }

  // When the user didn't map a broker cost column, skip the cost comparison
  // entirely instead of comparing against a defaulted 0 (which would flag
  // every position as a Warning with a misleading "cost off" note).
  const compareCost = opts.costCol >= 0;

  const out: PortfolioReconRow[] = [];
  for (let i = 0; i < allSymbols.length; i++) {
    const sym = allSymbols[i];
    const p = portfolioBySym[sym];
    const b = brokerBySym[sym];
    if (p && b) {
      const qtyDiff = p.quantity - b.qty;
      const costDiff = p.avgCostBasis - b.cost;
      const qtyBad = Math.abs(qtyDiff) > opts.qtyTolerance;
      const costBad = compareCost && Math.abs(costDiff) > opts.costTolerance;
      let status: PortfolioMatchStatus;
      let note: string;
      if (!qtyBad && !costBad) {
        status = "Matched";
        note = compareCost
          ? "Quantity and cost basis within tolerance"
          : "Quantity within tolerance (cost not compared)";
      } else {
        status = "Warning";
        const parts: string[] = [];
        if (qtyBad) parts.push(`qty off by ${qtyDiff.toFixed(4)}`);
        if (costBad) parts.push(`cost off by $${costDiff.toFixed(4)}`);
        note = parts.join("; ");
      }
      out.push({
        symbol: sym,
        portfolioQty: p.quantity,
        brokerQty: b.qty,
        qtyDiff,
        portfolioCost: p.avgCostBasis,
        brokerCost: compareCost ? b.cost : "",
        costDiff: compareCost ? costDiff : "",
        status,
        note,
      });
    } else if (p && !b) {
      // Portfolio-side cost is a known value regardless of whether the user
      // mapped a broker cost column — keep it visible.
      out.push({
        symbol: sym,
        portfolioQty: p.quantity,
        brokerQty: "",
        qtyDiff: "",
        portfolioCost: p.avgCostBasis,
        brokerCost: "",
        costDiff: "",
        status: "Error",
        note: "Missing from broker range",
      });
    } else if (!p && b) {
      out.push({
        symbol: sym,
        portfolioQty: "",
        brokerQty: b.qty,
        qtyDiff: "",
        portfolioCost: "",
        brokerCost: compareCost ? b.cost : "",
        costDiff: "",
        status: "Error",
        note: "Missing from portfolio",
      });
    }
  }
  return out;
}

export interface PortfolioReconSummary {
  matched: number;
  warnings: number;
  errors: number;
  total: number;
}

export function summarizePortfolioRecon(rows: PortfolioReconRow[]): PortfolioReconSummary {
  let matched = 0;
  let warnings = 0;
  let errors = 0;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].status === "Matched") matched++;
    else if (rows[i].status === "Warning") warnings++;
    else errors++;
  }
  return { matched, warnings, errors, total: rows.length };
}

const RECON_COLUMNS = [
  "Symbol",
  "Portfolio Qty",
  "Broker Qty",
  "Qty Difference",
  "Portfolio Cost Basis",
  "Broker Cost Basis",
  "Cost Difference",
  "Status",
  "Note",
];

const COLOR_MATCHED = "#DFF6DD";
const COLOR_WARNING = "#FFF4CE";
const COLOR_ERROR = "#FDE7E9";

/**
 * Write reconciliation output to a fresh worksheet. Returns the sheet name.
 */
export async function writePortfolioReconSheet(
  rows: PortfolioReconRow[],
  summary: PortfolioReconSummary,
  broker: BrokerRange
): Promise<string> {
  const baseName = sanitizeSheetName(
    `Portfolio Recon ${new Date().toISOString().slice(0, 10)}`
  );
  let finalName = baseName;

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

    // Summary block
    const summaryRows: Array<[string, string]> = [
      ["Portfolio Reconciliation", `Generated ${new Date().toLocaleString()}`],
      ["Broker Range", `${broker.address} (${broker.rowCount} rows × ${broker.columnCount} cols)`],
      [
        "Summary",
        `Matched: ${summary.matched}  ·  Warnings: ${summary.warnings}  ·  Errors: ${summary.errors}  ·  Total: ${summary.total}`,
      ],
    ];
    for (let r = 0; r < summaryRows.length; r++) {
      const range = sheet.getRangeByIndexes(r, 0, 1, 2);
      range.values = [[summaryRows[r][0], summaryRows[r][1]]];
      sheet.getRangeByIndexes(r, 0, 1, 1).format.font.bold = true;
      if (r === 0) {
        range.format.fill.color = "#0F6CBD";
        range.format.font.color = "#FFFFFF";
      }
    }

    const tableStart = summaryRows.length + 1;
    const headerRange = sheet.getRangeByIndexes(tableStart, 0, 1, RECON_COLUMNS.length);
    headerRange.values = [RECON_COLUMNS];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = COLOR_HEADER_FILL;

    if (rows.length > 0) {
      const dataStart = tableStart + 1;
      const grid: (string | number)[][] = new Array(rows.length);
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        grid[i] = [
          r.symbol,
          r.portfolioQty,
          r.brokerQty,
          r.qtyDiff,
          r.portfolioCost,
          r.brokerCost,
          r.costDiff,
          r.status,
          r.note,
        ];
      }
      const dataRange = sheet.getRangeByIndexes(dataStart, 0, rows.length, RECON_COLUMNS.length);
      dataRange.values = grid;

      // Numeric formats for qty (cols 1-3) and cost (cols 4-6)
      const numericCols: Array<{ col: number; format: string }> = [
        { col: 1, format: "#,##0.####" },
        { col: 2, format: "#,##0.####" },
        { col: 3, format: "#,##0.####;[Red]-#,##0.####" },
        { col: 4, format: "#,##0.0000" },
        { col: 5, format: "#,##0.0000" },
        { col: 6, format: "#,##0.0000;[Red]-#,##0.0000" },
      ];
      for (let k = 0; k < numericCols.length; k++) {
        const nc = numericCols[k];
        const col = sheet.getRangeByIndexes(dataStart, nc.col, rows.length, 1);
        const fmtGrid: string[][] = new Array(rows.length);
        for (let i = 0; i < rows.length; i++) fmtGrid[i] = [nc.format];
        col.numberFormat = fmtGrid;
      }

      // Conditional formatting by status
      const statusCol = 7;
      const statusLetter = columnIndexToLetter(statusCol);
      const rules: Array<{ status: PortfolioMatchStatus; color: string }> = [
        { status: "Matched", color: COLOR_MATCHED },
        { status: "Warning", color: COLOR_WARNING },
        { status: "Error", color: COLOR_ERROR },
      ];
      for (let i = 0; i < rules.length; i++) {
        const cf = dataRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = `=$${statusLetter}${dataStart + 1}="${rules[i].status}"`;
        cf.custom.format.fill.color = rules[i].color;
      }
    }

    const fullRange = sheet.getRangeByIndexes(
      0,
      0,
      tableStart + 1 + rows.length,
      RECON_COLUMNS.length
    );
    fullRange.format.autofitColumns();
    sheet.freezePanes.freezeRows(tableStart + 1);
    sheet.activate();
    await context.sync();
  });

  return finalName;
}

function sanitizeSheetName(name: string): string {
  const cleaned = name.replace(/[\\/?*[\]:]/g, "_").trim();
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned || "Sheet";
}
