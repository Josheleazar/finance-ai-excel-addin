/*
 * Reconciliation engine — pure, framework-free logic.
 *
 * Given two lists of normalized rows ("Source A" and "Source B") it produces a
 * set of 1:1 pairings plus leftovers using weighted fuzzy matching.
 *
 * Design notes:
 *   - Each candidate pair is scored on three dimensions: amount, date, description.
 *   - Each dimension yields a 0..1 score (1 = perfect match). Missing data on a
 *     dimension yields 0.5 (neutral) so it neither confirms nor penalizes.
 *   - Combined confidence = weighted average (amount 0.5, date 0.25, desc 0.25).
 *     Finance teams care about the amount first; date & description are corroborating.
 *   - Assignment is greedy: sort candidate pairs by confidence descending, take each
 *     pair iff both sides are still unmatched and confidence >= possibleThreshold.
 *     This is not optimal in the Hungarian sense but is O(N*M log(N*M)) and good
 *     enough for typical reconciliation sizes.
 *   - Remaining rows are reported as Unmatched.
 *
 * This module has no Office.js or DOM dependencies so it can be reused or unit tested.
 */

export type MatchingMode = "strict" | "normal" | "loose" | "custom";
export type MatchStatus = "Matched" | "Possible Match" | "Unmatched";

export interface Tolerances {
  /** Absolute amount tolerance in dollars (or whatever currency). */
  amountFixed: number;
  /** Amount tolerance as a percentage of the larger of the two amounts (0..100). */
  amountPct: number;
  /** Whole-day tolerance applied to the absolute date delta. */
  dateDays: number;
}

/** Confidence thresholds that map a pair's score to a MatchStatus. */
export interface Thresholds {
  /** Confidence >= match => "Matched". */
  match: number;
  /** possible <= confidence < match => "Possible Match"; below possible => not paired. */
  possible: number;
}

/** Preset tolerances per matching mode. `custom` stores the last user-entered values
 *  at init time and is otherwise just a "don't apply a preset" sentinel. */
export const DEFAULT_TOLERANCES: { [K in MatchingMode]: Tolerances } = {
  strict: { amountFixed: 0, amountPct: 0, dateDays: 0 },
  normal: { amountFixed: 0.01, amountPct: 0, dateDays: 1 },
  loose: { amountFixed: 1.0, amountPct: 1.0, dateDays: 3 },
  custom: { amountFixed: 0.01, amountPct: 0, dateDays: 1 },
};

export const DEFAULT_THRESHOLDS: { [K in MatchingMode]: Thresholds } = {
  strict: { match: 0.95, possible: 0.85 },
  normal: { match: 0.85, possible: 0.7 },
  loose: { match: 0.7, possible: 0.55 },
  custom: { match: 0.85, possible: 0.7 },
};

export interface NormalizedRow {
  /** Zero-based position within the source data rows (excludes header). */
  index: number;
  /** 1-based Excel row number in the source worksheet, for user-friendly refs. */
  excelRow: number;
  /** Parsed date as ms since epoch, or null if unparseable. */
  dateMs: number | null;
  /** Parsed amount, or null if unparseable. */
  amount: number | null;
  /** Lowercased, punctuation-stripped description for fuzzy compare. */
  descriptionNorm: string;
  /** Original-ish display values preserved for output. */
  dateDisplay: string;
  amountDisplay: number | string;
  descriptionDisplay: string;
}

export interface Match {
  a: NormalizedRow;
  b: NormalizedRow;
  confidence: number;
  status: MatchStatus;
  amountDiff: number;
  dateDiffDays: number;
  descSimilarity: number;
  notes: string;
}

export interface ReconciliationResult {
  matches: Match[];
  unmatchedA: NormalizedRow[];
  unmatchedB: NormalizedRow[];
  summary: {
    matched: number;
    possible: number;
    unmatchedA: number;
    unmatchedB: number;
    totalA: number;
    totalB: number;
  };
}

/* ------------------------------- Normalizers ------------------------------ */

// Excel stores dates as serial days since 1899-12-30 (accounting for the 1900
// leap-year bug). 25569 days separate 1899-12-30 from the Unix epoch.
const EXCEL_EPOCH_DAYS_OFFSET = 25569;
const MS_PER_DAY = 86400 * 1000;

/**
 * Parse an Excel cell value as a date, returning ms since epoch or null.
 * Accepts Excel serial numbers, JS Date instances, and ISO/locale date strings.
 */
export function parseDateCell(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    const t = value.getTime();
    return isNaN(t) ? null : t;
  }
  if (typeof value === "number") {
    // Valid Excel date serials are roughly 1..2958465 (0001-01-01 .. 9999-12-31).
    if (value < 1 || value > 2958465) return null;
    return Math.round((value - EXCEL_EPOCH_DAYS_OFFSET) * MS_PER_DAY);
  }
  if (typeof value === "string") {
    const s = value.trim();
    if (!s) return null;
    const t = Date.parse(s);
    return isNaN(t) ? null : t;
  }
  return null;
}

/**
 * Parse an Excel cell value as a number. Tolerates common finance formatting:
 * "$1,234.56", "(123.45)" (accounting negatives), "1.234,56" (European style).
 */
export function parseAmountCell(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value !== "string") return null;

  let s = value.trim();
  if (!s) return null;

  // Accounting-style negatives: "(123.45)" => -123.45
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }

  // Drop currency symbols and whitespace.
  s = s.replace(/[\s$€£¥]/g, "");

  // Disambiguate decimal separator when both "," and "." are present.
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) {
      // European: "1.234,56"
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // US: "1,234.56"
      s = s.replace(/,/g, "");
    }
  } else if (lastComma >= 0) {
    // Only commas. If there's a plausible decimal comma (1-2 digits after the
    // last comma), treat it as decimal; otherwise as thousands separators.
    const tail = s.length - lastComma - 1;
    if (tail === 1 || tail === 2) {
      s = s.replace(/,/g, ".");
      // If there were multiple commas, keep only the last as the decimal.
      const firstDot = s.indexOf(".");
      const lastDot2 = s.lastIndexOf(".");
      if (firstDot !== lastDot2) {
        s = s.slice(0, lastDot2).replace(/\./g, "") + s.slice(lastDot2);
      }
    } else {
      s = s.replace(/,/g, "");
    }
  }

  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return negative ? -n : n;
}

/** Lowercase, collapse punctuation to single spaces, trim. */
export function normalizeDescription(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/* --------------------------- Levenshtein similarity --------------------------- */

/** Classic Levenshtein edit distance. O(|a|*|b|) time, O(min(|a|,|b|)) space. */
export function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  if (a.length > b.length) {
    const t = a;
    a = b;
    b = t;
  }
  const prev = new Array<number>(a.length + 1);
  const curr = new Array<number>(a.length + 1);
  for (let i = 0; i <= a.length; i++) prev[i] = i;
  for (let j = 1; j <= b.length; j++) {
    curr[0] = j;
    for (let i = 1; i <= a.length; i++) {
      const cost = a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1;
      curr[i] = Math.min(curr[i - 1] + 1, prev[i] + 1, prev[i - 1] + cost);
    }
    for (let i = 0; i <= a.length; i++) prev[i] = curr[i];
  }
  return prev[a.length];
}

/** 1.0 = identical; 0.0 = completely different. Both empty strings => 1.0. */
export function stringSimilarity(a: string, b: string): number {
  if (!a && !b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

/* -------------------------- Per-dimension scoring -------------------------- */

function amountScore(a: number | null, b: number | null, tol: Tolerances): number {
  if (a === null || b === null) return 0.5;
  const diff = Math.abs(a - b);
  const magnitude = Math.max(Math.abs(a), Math.abs(b), 1);
  const allowed = Math.max(tol.amountFixed, (tol.amountPct / 100) * magnitude);
  if (diff <= allowed + 1e-9) return 1;
  // Linear decay: reaches 0 when the excess equals the magnitude itself.
  return Math.max(0, 1 - (diff - allowed) / magnitude);
}

function dateScore(aMs: number | null, bMs: number | null, tol: Tolerances): number {
  if (aMs === null || bMs === null) return 0.5;
  const days = Math.abs(aMs - bMs) / MS_PER_DAY;
  if (days <= tol.dateDays) return 1;
  // Linear decay: fully penalized 30 days beyond the tolerance window.
  return Math.max(0, 1 - (days - tol.dateDays) / 30);
}

function descriptionScore(a: string, b: string): number {
  if (!a && !b) return 0.5;
  if (!a || !b) return 0.3;
  return stringSimilarity(a, b);
}

/** Compute the weighted confidence of pairing (a,b) plus its sub-scores. */
export function pairConfidence(
  a: NormalizedRow,
  b: NormalizedRow,
  tol: Tolerances
): { confidence: number; amountS: number; dateS: number; descS: number } {
  const amountS = amountScore(a.amount, b.amount, tol);
  const dateS = dateScore(a.dateMs, b.dateMs, tol);
  const descS = descriptionScore(a.descriptionNorm, b.descriptionNorm);
  const confidence = amountS * 0.5 + dateS * 0.25 + descS * 0.25;
  return { confidence, amountS, dateS, descS };
}

/* ------------------------------- Matching ------------------------------- */

/**
 * Greedy 1:1 matching: compute confidences for all candidate pairs, sort desc,
 * accept in order as long as both sides remain free and we're above the
 * possible-match threshold.
 */
export function reconcile(
  sourceA: NormalizedRow[],
  sourceB: NormalizedRow[],
  tol: Tolerances,
  thresholds: Thresholds
): ReconciliationResult {
  interface Candidate {
    i: number;
    j: number;
    confidence: number;
    amountS: number;
    dateS: number;
    descS: number;
  }

  const candidates: Candidate[] = [];
  for (let i = 0; i < sourceA.length; i++) {
    const a = sourceA[i];
    for (let j = 0; j < sourceB.length; j++) {
      const b = sourceB[j];
      const pc = pairConfidence(a, b, tol);
      if (pc.confidence >= thresholds.possible) {
        candidates.push({
          i,
          j,
          confidence: pc.confidence,
          amountS: pc.amountS,
          dateS: pc.dateS,
          descS: pc.descS,
        });
      }
    }
  }
  candidates.sort((x, y) => y.confidence - x.confidence);

  const usedA: boolean[] = new Array(sourceA.length);
  const usedB: boolean[] = new Array(sourceB.length);
  const matches: Match[] = [];

  for (let k = 0; k < candidates.length; k++) {
    const c = candidates[k];
    if (usedA[c.i] || usedB[c.j]) continue;
    const a = sourceA[c.i];
    const b = sourceB[c.j];
    const status: MatchStatus = c.confidence >= thresholds.match ? "Matched" : "Possible Match";
    const amountDiff = (a.amount === null ? 0 : a.amount) - (b.amount === null ? 0 : b.amount);
    const dateDiffDays =
      a.dateMs !== null && b.dateMs !== null ? Math.round((a.dateMs - b.dateMs) / MS_PER_DAY) : 0;
    matches.push({
      a,
      b,
      confidence: c.confidence,
      status,
      amountDiff,
      dateDiffDays,
      descSimilarity: c.descS,
      notes: buildMatchNotes(c.amountS, c.dateS, c.descS, amountDiff, dateDiffDays),
    });
    usedA[c.i] = true;
    usedB[c.j] = true;
  }

  const unmatchedA: NormalizedRow[] = [];
  for (let i = 0; i < sourceA.length; i++) if (!usedA[i]) unmatchedA.push(sourceA[i]);
  const unmatchedB: NormalizedRow[] = [];
  for (let j = 0; j < sourceB.length; j++) if (!usedB[j]) unmatchedB.push(sourceB[j]);

  let matched = 0;
  let possible = 0;
  for (let k = 0; k < matches.length; k++) {
    if (matches[k].status === "Matched") matched++;
    else possible++;
  }

  return {
    matches,
    unmatchedA,
    unmatchedB,
    summary: {
      matched,
      possible,
      unmatchedA: unmatchedA.length,
      unmatchedB: unmatchedB.length,
      totalA: sourceA.length,
      totalB: sourceB.length,
    },
  };
}

function buildMatchNotes(
  _amountS: number,
  _dateS: number,
  descS: number,
  amountDiff: number,
  dateDiffDays: number
): string {
  const parts: string[] = [];
  if (Math.abs(amountDiff) > 1e-9) {
    const sign = amountDiff >= 0 ? "+" : "-";
    parts.push(`Amount ${sign}${Math.abs(amountDiff).toFixed(2)}`);
  }
  if (dateDiffDays !== 0) {
    const sign = dateDiffDays > 0 ? "+" : "";
    const unit = Math.abs(dateDiffDays) === 1 ? "day" : "days";
    parts.push(`${sign}${dateDiffDays} ${unit}`);
  }
  if (descS < 1) {
    parts.push(`Desc ${Math.round(descS * 100)}%`);
  }
  if (!parts.length) parts.push("Exact match");
  return parts.join(" · ");
}
