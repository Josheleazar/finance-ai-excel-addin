/*
 * Reconciliation engine — dynamic field-based matching.
 *
 * A reconciliation is defined by 1..N user-chosen fields. Each field picks
 * one column from Source A and one from Source B and a match type:
 *
 *   - exact   : case-insensitive string equality (toggleable case-sensitive).
 *   - numeric : fixed $ + percentage tolerance; linear decay past the window.
 *   - date    : day-level tolerance; linear decay past the window.
 *   - fuzzy   : Levenshtein similarity 0..1 with a pass threshold.
 *
 * Each field carries a weight (low/medium/high → 1/2/3) and an optional
 * "required" flag. When a required field fails its pass test the pair is
 * either excluded entirely (hard gate, default) or capped to "Possible Match"
 * (downgrade, opt-in).
 *
 * Confidence = weighted average of per-field sub-scores.
 * Pairs above matchThreshold → "Matched" (unless capped).
 * Pairs above possibleThreshold → "Possible Match".
 * Pairs below possibleThreshold → not paired (row stays unmatched).
 *
 * Assignment is greedy 1:1 on descending confidence. This file has no
 * Office.js or DOM dependencies so it can be unit tested in isolation.
 */

export type MatchStatus = "Matched" | "Possible Match" | "Unmatched";
export type FieldType = "exact" | "numeric" | "date" | "fuzzy";
export type FieldWeight = "low" | "medium" | "high";
export type Sensitivity = "strict" | "normal" | "loose";

export const MAX_FIELDS = 6;

export interface FieldTolerance {
  /** numeric: absolute tolerance in whatever units the column uses. */
  amountFixed?: number;
  /** numeric: percentage tolerance (0..100) applied to max(|a|, |b|, 1). */
  amountPct?: number;
  /** date: tolerance in whole days. */
  dateDays?: number;
  /** fuzzy: minimum Levenshtein similarity (0..1) to count as a "pass". */
  minSimilarity?: number;
  /** exact: when true, string comparison is case-sensitive. */
  caseSensitive?: boolean;
}

export interface FieldConfig {
  /** Stable id used by the UI to key DOM rows. */
  id: string;
  /** User-facing label (used in output headers and notes). */
  label: string;
  type: FieldType;
  /** Column index in Source A (0-based). -1 = not assigned. */
  colA: number;
  /** Column index in Source B (0-based). -1 = not assigned. */
  colB: number;
  weight: FieldWeight;
  required: boolean;
  /** If required & this field fails: true = cap to Possible, false = exclude pair. */
  downgradeOnFail: boolean;
  tolerance: FieldTolerance;
}

export interface Thresholds {
  match: number;
  possible: number;
}

export const SENSITIVITY_THRESHOLDS: { [K in Sensitivity]: Thresholds } = {
  strict: { match: 0.95, possible: 0.85 },
  normal: { match: 0.85, possible: 0.7 },
  loose: { match: 0.7, possible: 0.55 },
};

export const WEIGHT_VALUES: { [K in FieldWeight]: number } = {
  low: 1,
  medium: 2,
  high: 3,
};

export function defaultToleranceForType(type: FieldType): FieldTolerance {
  switch (type) {
    case "exact":
      return { caseSensitive: false };
    case "numeric":
      return { amountFixed: 0.01, amountPct: 0 };
    case "date":
      return { dateDays: 1 };
    case "fuzzy":
      return { minSimilarity: 0.6 };
  }
}

export interface RawRow {
  /** 1-based Excel row number in the source worksheet, for user-friendly refs. */
  excelRow: number;
  /** 0-based index within the data rows (post header trim, post blank skip). */
  index: number;
  /** Full row of raw cell values, indexable by column. */
  cells: unknown[];
}

export interface FieldScore {
  /** 0..1 sub-score contributed to confidence. */
  score: number;
  /** Whether the field "passes" its pass test (used for required-field gating). */
  passed: boolean;
}

export interface PairScore {
  /** If true, a required hard-gate field failed — skip this pair entirely. */
  excluded: boolean;
  /** If true, cap the final status at "Possible Match" regardless of confidence. */
  capAsPossible: boolean;
  /** Weighted average of per-field sub-scores. 0..1. */
  confidence: number;
  perField: FieldScore[];
}

export interface Match {
  a: RawRow;
  b: RawRow;
  confidence: number;
  status: MatchStatus;
  perField: FieldScore[];
  notes: string;
}

export interface ReconciliationResult {
  matches: Match[];
  unmatchedA: RawRow[];
  unmatchedB: RawRow[];
  summary: {
    matched: number;
    possible: number;
    unmatchedA: number;
    unmatchedB: number;
    totalA: number;
    totalB: number;
  };
}

/* ------------------------------- Parsers ------------------------------- */

// Excel stores dates as serial days since 1899-12-30 (accounting for the 1900
// leap-year bug). 25569 days separate 1899-12-30 from the Unix epoch.
const EXCEL_EPOCH_DAYS_OFFSET = 25569;
const MS_PER_DAY = 86400 * 1000;

export function parseDateCell(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    const t = value.getTime();
    return isNaN(t) ? null : t;
  }
  if (typeof value === "number") {
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

  s = s.replace(/[\s$€£¥]/g, "");

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
    const tail = s.length - lastComma - 1;
    if (tail === 1 || tail === 2) {
      s = s.replace(/,/g, ".");
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

export function normalizeDescription(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/* --------------------------- Levenshtein ---------------------------- */

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

export function stringSimilarity(a: string, b: string): number {
  if (!a && !b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

/* ------------------------- Per-field scoring ------------------------- */

function coerce(value: number | undefined, fallback: number): number {
  return typeof value === "number" && isFinite(value) ? value : fallback;
}

/**
 * Score one field for a candidate (a, b) pair. Returns both a 0..1 sub-score
 * and a `passed` flag used by required-field gating.
 */
export function scoreField(field: FieldConfig, valA: unknown, valB: unknown): FieldScore {
  switch (field.type) {
    case "exact": {
      const caseSensitive = !!field.tolerance.caseSensitive;
      const sa = valA === null || valA === undefined ? "" : String(valA).trim();
      const sb = valB === null || valB === undefined ? "" : String(valB).trim();
      if (!sa || !sb) return { score: 0, passed: false };
      const eq = caseSensitive ? sa === sb : sa.toLowerCase() === sb.toLowerCase();
      return eq ? { score: 1, passed: true } : { score: 0, passed: false };
    }
    case "numeric": {
      const a = parseAmountCell(valA);
      const b = parseAmountCell(valB);
      if (a === null || b === null) return { score: 0, passed: false };
      const fixed = coerce(field.tolerance.amountFixed, 0);
      const pct = coerce(field.tolerance.amountPct, 0);
      const magnitude = Math.max(Math.abs(a), Math.abs(b), 1);
      const allowed = Math.max(fixed, (pct / 100) * magnitude);
      const diff = Math.abs(a - b);
      if (diff <= allowed + 1e-9) return { score: 1, passed: true };
      return { score: Math.max(0, 1 - (diff - allowed) / magnitude), passed: false };
    }
    case "date": {
      const a = parseDateCell(valA);
      const b = parseDateCell(valB);
      if (a === null || b === null) return { score: 0, passed: false };
      const days = Math.abs(a - b) / MS_PER_DAY;
      const tol = coerce(field.tolerance.dateDays, 0);
      if (days <= tol) return { score: 1, passed: true };
      // Linear decay: fully penalized 30 days beyond the tolerance window.
      return { score: Math.max(0, 1 - (days - tol) / 30), passed: false };
    }
    case "fuzzy": {
      const a = normalizeDescription(valA);
      const b = normalizeDescription(valB);
      // Treat any missing value as a hard 0 — otherwise a blank cell would
      // bleed non-zero confidence into the weighted average for what is
      // effectively a missing field.
      if (!a || !b) return { score: 0, passed: false };
      const sim = stringSimilarity(a, b);
      const threshold = coerce(field.tolerance.minSimilarity, 0.6);
      return { score: sim, passed: sim >= threshold };
    }
  }
}

/**
 * Score a candidate pair across all fields, applying required-field gating.
 */
export function scorePair(fields: FieldConfig[], rowA: RawRow, rowB: RawRow): PairScore {
  const perField: FieldScore[] = new Array(fields.length);
  let excluded = false;
  let capAsPossible = false;
  let weightedSum = 0;
  let weightSum = 0;

  for (let k = 0; k < fields.length; k++) {
    const f = fields[k];
    const valA = f.colA >= 0 ? rowA.cells[f.colA] : undefined;
    const valB = f.colB >= 0 ? rowB.cells[f.colB] : undefined;
    const s = scoreField(f, valA, valB);
    perField[k] = s;

    if (f.required && !s.passed) {
      if (f.downgradeOnFail) capAsPossible = true;
      else excluded = true;
    }

    const w = WEIGHT_VALUES[f.weight];
    weightedSum += s.score * w;
    weightSum += w;
  }

  const confidence = weightSum > 0 ? weightedSum / weightSum : 0;
  return { excluded, capAsPossible, confidence, perField };
}

/* ------------------------------ Matching ------------------------------ */

/**
 * Greedy 1:1 matching: compute confidence for every non-excluded candidate
 * pair at or above the possible threshold, sort desc, accept in order as
 * long as both sides remain free.
 */
export function reconcile(
  sourceA: RawRow[],
  sourceB: RawRow[],
  fields: FieldConfig[],
  thresholds: Thresholds
): ReconciliationResult {
  interface Candidate {
    i: number;
    j: number;
    confidence: number;
    capAsPossible: boolean;
    perField: FieldScore[];
  }

  const candidates: Candidate[] = [];
  for (let i = 0; i < sourceA.length; i++) {
    for (let j = 0; j < sourceB.length; j++) {
      const ps = scorePair(fields, sourceA[i], sourceB[j]);
      if (ps.excluded) continue;
      if (ps.confidence < thresholds.possible) continue;
      candidates.push({
        i,
        j,
        confidence: ps.confidence,
        capAsPossible: ps.capAsPossible,
        perField: ps.perField,
      });
    }
  }
  candidates.sort((x, y) => y.confidence - x.confidence);

  const usedA: boolean[] = new Array(sourceA.length);
  const usedB: boolean[] = new Array(sourceB.length);
  const matches: Match[] = [];

  for (let k = 0; k < candidates.length; k++) {
    const c = candidates[k];
    if (usedA[c.i] || usedB[c.j]) continue;
    const isMatched = c.confidence >= thresholds.match && !c.capAsPossible;
    const status: MatchStatus = isMatched ? "Matched" : "Possible Match";
    matches.push({
      a: sourceA[c.i],
      b: sourceB[c.j],
      confidence: c.confidence,
      status,
      perField: c.perField,
      notes: buildNotes(fields, c.perField, c.capAsPossible),
    });
    usedA[c.i] = true;
    usedB[c.j] = true;
  }

  const unmatchedA: RawRow[] = [];
  for (let i = 0; i < sourceA.length; i++) if (!usedA[i]) unmatchedA.push(sourceA[i]);
  const unmatchedB: RawRow[] = [];
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

function buildNotes(fields: FieldConfig[], perField: FieldScore[], capped: boolean): string {
  const parts: string[] = [];
  if (capped) parts.push("Capped to Possible (required field mismatch)");
  for (let k = 0; k < fields.length; k++) {
    const f = fields[k];
    const s = perField[k];
    if (s.score < 1) {
      const tag = f.required && !s.passed ? " ⚠" : "";
      parts.push(`${f.label}: ${Math.round(s.score * 100)}%${tag}`);
    }
  }
  if (!parts.length) parts.push("All fields perfect");
  return parts.join(" · ");
}
