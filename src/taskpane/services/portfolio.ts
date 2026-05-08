/*
 * Portfolio tracker — pure domain logic.
 *
 * Responsibilities
 *   - Define the Holding shape used throughout the portfolio feature.
 *   - Compute per-row metrics (market value, unrealized P&L $/%, weight %).
 *   - Parse broker CSVs and suggest a column-to-field mapping ("smart mapping").
 *
 * This module is deliberately free of Office.js / DOM dependencies so it can
 * be exercised in unit tests and reused by both the task pane UI and the
 * Excel I/O layer.
 */

import { parseAmountCell } from "./reconciliation";

/** Canonical fields a broker CSV can map onto. Order matters for the UI. */
export type HoldingField =
  | "symbol"
  | "quantity"
  | "avgCostBasis"
  | "purchaseDate"
  | "accountType";

export const HOLDING_FIELDS: readonly HoldingField[] = [
  "symbol",
  "quantity",
  "avgCostBasis",
  "purchaseDate",
  "accountType",
] as const;

export const HOLDING_FIELD_LABELS: { [K in HoldingField]: string } = {
  symbol: "Symbol",
  quantity: "Quantity",
  avgCostBasis: "Avg Cost Basis",
  purchaseDate: "Purchase Date",
  accountType: "Account Type",
};

/** Which import fields are required; the rest are optional metadata. */
export const REQUIRED_IMPORT_FIELDS: readonly HoldingField[] = [
  "symbol",
  "quantity",
  "avgCostBasis",
] as const;

/** A single portfolio row. Price + companyName are enriched from Finnhub. */
export interface Holding {
  /** Ticker, stored uppercase. The unique key for a row. */
  symbol: string;
  /** Optional company name (filled from Finnhub profile). */
  companyName?: string;
  /** Optional sector / industry (filled from Finnhub profile). */
  sector?: string;
  /** Number of shares. */
  quantity: number;
  /** Per-share average cost basis in the account's currency. */
  avgCostBasis: number;
  /** Last known market price from Finnhub. Empty until refreshed. */
  currentPrice?: number;
  /** ISO date string (YYYY-MM-DD) when the position was opened — optional. */
  purchaseDate?: string;
  /** Free-form account bucket (e.g. "Taxable", "IRA"). Optional. */
  accountType?: string;
  /** Unix ms timestamp of the last successful price refresh. */
  lastRefreshedAt?: number;
}

/** Per-row derived metrics + the total-portfolio context needed to compute weight %. */
export interface HoldingMetrics {
  marketValue: number;
  costTotal: number;
  unrealizedPnl: number;
  unrealizedPnlPct: number;
  /** 0..100 share of the total market value. */
  weightPct: number;
}

export interface PortfolioTotals {
  marketValue: number;
  costTotal: number;
  unrealizedPnl: number;
  unrealizedPnlPct: number;
  rowCount: number;
}

/* ----------------------------- Math ----------------------------- */

/** Compute market value, cost total, and unrealized P&L for a single holding.
 *  `totalMarketValue` is used only to compute weight %; pass 0 if unknown. */
export function computeHoldingMetrics(
  h: Holding,
  totalMarketValue: number
): HoldingMetrics {
  const qty = toFiniteNumber(h.quantity);
  const cost = toFiniteNumber(h.avgCostBasis);
  const price = toFiniteNumber(h.currentPrice);
  const marketValue = qty * price;
  const costTotal = qty * cost;
  const unrealizedPnl = marketValue - costTotal;
  const unrealizedPnlPct = costTotal > 0 ? (unrealizedPnl / costTotal) * 100 : 0;
  const weightPct =
    totalMarketValue > 0 ? (marketValue / totalMarketValue) * 100 : 0;
  return { marketValue, costTotal, unrealizedPnl, unrealizedPnlPct, weightPct };
}

/** Sum market value / cost across all holdings. */
export function computeTotals(holdings: Holding[]): PortfolioTotals {
  let marketValue = 0;
  let costTotal = 0;
  for (let i = 0; i < holdings.length; i++) {
    const qty = toFiniteNumber(holdings[i].quantity);
    const price = toFiniteNumber(holdings[i].currentPrice);
    const cost = toFiniteNumber(holdings[i].avgCostBasis);
    marketValue += qty * price;
    costTotal += qty * cost;
  }
  const unrealizedPnl = marketValue - costTotal;
  const unrealizedPnlPct = costTotal > 0 ? (unrealizedPnl / costTotal) * 100 : 0;
  return {
    marketValue,
    costTotal,
    unrealizedPnl,
    unrealizedPnlPct,
    rowCount: holdings.length,
  };
}

function toFiniteNumber(n: unknown): number {
  return typeof n === "number" && isFinite(n) ? n : 0;
}

/** Normalise a ticker to uppercase trimmed form. Returns "" if unusable. */
export function normalizeSymbol(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  return String(raw).trim().toUpperCase();
}

/* --------------------------- CSV parsing --------------------------- */

/** Row-level CSV parse output before field mapping is applied. */
export interface ParsedCsv {
  headers: string[];
  rows: string[][];
}

/**
 * Parse a CSV string into headers + rows.
 *
 * Handles the common real-world cases:
 *   - quoted fields with embedded commas or newlines (RFC 4180 style)
 *   - escaped quotes inside quoted fields ("" → ")
 *   - \r\n, \n, and \r line endings
 *   - trailing blank lines
 *
 * Does NOT try to infer a delimiter; comma is assumed. (Most broker exports
 * are comma-delimited; TSV support can be added later.)
 */
export function parseCsv(text: string): ParsedCsv {
  const rows: string[][] = [];
  let cur: string[] = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    if (inQuotes) {
      if (ch === '"') {
        if (text.charAt(i + 1) === '"') {
          field += '"';
          i++; // consume the escaped quote
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
      cur.push(field);
      field = "";
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      // collapse \r\n into a single terminator
      if (ch === "\r" && text.charAt(i + 1) === "\n") i++;
      cur.push(field);
      field = "";
      rows.push(cur);
      cur = [];
      continue;
    }
    field += ch;
  }
  // flush
  if (field.length > 0 || cur.length > 0) {
    cur.push(field);
    rows.push(cur);
  }

  // Drop fully-blank rows (common trailing whitespace in exports).
  const cleaned: string[][] = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    let hasContent = false;
    for (let c = 0; c < row.length; c++) {
      if (row[c] && row[c].trim()) {
        hasContent = true;
        break;
      }
    }
    if (hasContent) cleaned.push(row);
  }

  if (!cleaned.length) return { headers: [], rows: [] };
  const headers = cleaned[0].map((h) => h.trim());
  const body = cleaned.slice(1);
  return { headers, rows: body };
}

/* ----------------------- Smart column mapping ----------------------- */

/**
 * Regexes the "smart mapping" uses to guess which CSV column maps to which
 * standard field. Tuned for Schwab / Fidelity / Vanguard / Robinhood / IBKR
 * exports — add patterns here rather than special-casing callers.
 */
const HEADER_HINTS: { [K in HoldingField]: RegExp[] } = {
  symbol: [/^symbol$/i, /ticker/i, /security.*id/i, /^sym$/i, /cusip/i],
  quantity: [
    /^qty$/i,
    /quantity/i,
    /shares?/i,
    /units?/i,
    /position/i,
    /holding/i,
  ],
  avgCostBasis: [
    /avg.*cost/i,
    /average.*cost/i,
    /cost.*basis/i,
    /cost.*per.*share/i,
    /purchase.*price/i,
    /buy.*price/i,
    /unit.*cost/i,
  ],
  purchaseDate: [
    /purchase.*date/i,
    /trade.*date/i,
    /acquired/i,
    /opened/i,
    /date/i,
  ],
  accountType: [/account.*type/i, /account/i, /^type$/i, /registration/i],
};

export interface ColumnMapping {
  /** Mapping from canonical field → CSV column index. -1 means "unmapped". */
  [K: string]: number;
}

/**
 * Suggest a column mapping by scanning the CSV headers.
 *
 * Strategy:
 *   1. For each canonical field in priority order, try each regex hint in
 *      order. First unused header column that matches wins.
 *   2. Already-claimed columns are skipped so fields don't collide.
 *   3. Fields with no match get -1 so the user can pick (or leave unmapped).
 */
export function suggestColumnMapping(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {
    symbol: -1,
    quantity: -1,
    avgCostBasis: -1,
    purchaseDate: -1,
    accountType: -1,
  };
  const taken: { [col: number]: boolean } = {};
  for (let f = 0; f < HOLDING_FIELDS.length; f++) {
    const field = HOLDING_FIELDS[f];
    const hints = HEADER_HINTS[field];
    for (let h = 0; h < hints.length && mapping[field] < 0; h++) {
      for (let c = 0; c < headers.length; c++) {
        if (taken[c]) continue;
        if (hints[h].test(headers[c])) {
          mapping[field] = c;
          taken[c] = true;
          break;
        }
      }
    }
  }
  return mapping;
}

/**
 * Apply a column mapping to parsed CSV rows and return Holdings.
 *
 * Invalid rows (missing symbol, non-numeric quantity, etc.) are collected
 * into `errors` so the UI can show a rollup summary instead of silently
 * dropping data.
 */
export interface ImportResult {
  holdings: Holding[];
  errors: string[];
  /** Count of rows skipped due to validation failures. */
  skipped: number;
}

export function applyMapping(parsed: ParsedCsv, mapping: ColumnMapping): ImportResult {
  const holdings: Holding[] = [];
  const errors: string[] = [];
  let skipped = 0;

  // Require symbol / quantity / avgCostBasis; other fields are optional.
  const missing: HoldingField[] = [];
  for (let i = 0; i < REQUIRED_IMPORT_FIELDS.length; i++) {
    const f = REQUIRED_IMPORT_FIELDS[i];
    if (mapping[f] === undefined || mapping[f] < 0) missing.push(f);
  }
  if (missing.length) {
    errors.push(
      `Missing required column mapping: ${missing.map((f) => HOLDING_FIELD_LABELS[f]).join(", ")}.`
    );
    return { holdings, errors, skipped: parsed.rows.length };
  }

  for (let r = 0; r < parsed.rows.length; r++) {
    const row = parsed.rows[r];
    const rowNum = r + 2; // +1 for 1-based, +1 for header row
    const symbol = normalizeSymbol(row[mapping.symbol]);
    if (!symbol) {
      skipped++;
      errors.push(`Row ${rowNum}: missing symbol.`);
      continue;
    }
    const qty = parseLooseNumber(row[mapping.quantity]);
    if (qty === null || qty <= 0) {
      skipped++;
      errors.push(`Row ${rowNum} (${symbol}): invalid quantity "${row[mapping.quantity] || ""}".`);
      continue;
    }
    const cost = parseLooseNumber(row[mapping.avgCostBasis]);
    if (cost === null || cost < 0) {
      skipped++;
      errors.push(`Row ${rowNum} (${symbol}): invalid cost basis "${row[mapping.avgCostBasis] || ""}".`);
      continue;
    }
    const holding: Holding = {
      symbol,
      quantity: qty,
      avgCostBasis: cost,
    };
    if (mapping.purchaseDate >= 0) {
      const d = normalizeDateString(row[mapping.purchaseDate]);
      if (d) holding.purchaseDate = d;
    }
    if (mapping.accountType >= 0) {
      const v = (row[mapping.accountType] || "").trim();
      if (v) holding.accountType = v;
    }
    holdings.push(holding);
  }

  return { holdings, errors, skipped };
}

/**
 * Loose number parser that accepts typical broker-export formatting:
 *   "1,234.56", "$1,234.56", "(1,234.56)" (accounting negative), "1.234,56"
 * Returns null for empty / unparseable values.
 *
 * Thin re-export of `parseAmountCell` from the reconciliation service so both
 * features use identical numeric-parsing behavior.
 */
export function parseLooseNumber(raw: unknown): number | null {
  return parseAmountCell(raw);
}

/** Try to turn a raw cell into an ISO date string (YYYY-MM-DD). */
export function normalizeDateString(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  const s = String(raw).trim();
  if (!s) return "";
  const t = Date.parse(s);
  if (!isNaN(t)) {
    return new Date(t).toISOString().slice(0, 10);
  }
  return s; // fall back to whatever the broker supplied
}

/* ---------------------- Merge helpers ---------------------- */

/**
 * Merge incoming holdings into an existing list. When a symbol already exists,
 * the incoming row's quantity and cost basis replace the existing values
 * (simpler than recomputing a weighted average — per the requirements we do
 * not implement full transaction history cost-basis calculation).
 *
 * Enriched fields (companyName, sector, currentPrice) are preserved from the
 * existing row so a fresh import doesn't blow away prior price refreshes.
 */
export function mergeHoldings(existing: Holding[], incoming: Holding[]): Holding[] {
  const bySymbol: { [sym: string]: Holding } = {};
  for (let i = 0; i < existing.length; i++) {
    bySymbol[existing[i].symbol] = { ...existing[i] };
  }
  for (let i = 0; i < incoming.length; i++) {
    const row = incoming[i];
    const prev = bySymbol[row.symbol];
    if (prev) {
      bySymbol[row.symbol] = {
        ...prev,
        quantity: row.quantity,
        avgCostBasis: row.avgCostBasis,
        purchaseDate: row.purchaseDate || prev.purchaseDate,
        accountType: row.accountType || prev.accountType,
      };
    } else {
      bySymbol[row.symbol] = { ...row };
    }
  }
  // Preserve insertion order: existing first (in original order), then new symbols.
  const out: Holding[] = [];
  const seen: { [s: string]: boolean } = {};
  for (let i = 0; i < existing.length; i++) {
    const s = existing[i].symbol;
    if (bySymbol[s] && !seen[s]) {
      out.push(bySymbol[s]);
      seen[s] = true;
    }
  }
  for (let i = 0; i < incoming.length; i++) {
    const s = incoming[i].symbol;
    if (bySymbol[s] && !seen[s]) {
      out.push(bySymbol[s]);
      seen[s] = true;
    }
  }
  return out;
}
