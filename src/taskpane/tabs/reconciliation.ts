/*
 * Reconciliation tab — orchestrates the end-to-end user flow:
 *
 *   1. User selects a range in Excel and clicks "Use selection" for Source A or B.
 *   2. We auto-detect whether the first row is a header and which columns map to
 *      Date / Amount / Description; the user can override via dropdowns.
 *   3. User picks a matching mode preset (Strict / Normal / Loose / Custom) and
 *      fine-tunes amount & date tolerances.
 *   4. "Run Reconciliation" reads both ranges, runs the fuzzy matcher, and writes
 *      a results worksheet with a summary block and color-coded rows.
 *   5. "Save as Default" persists the current tolerance configuration.
 *
 * All Office.js calls live in services/reconcileExcel; all matching logic lives
 * in services/reconciliation. This file is the glue (DOM <-> services).
 */

/* global HTMLButtonElement, HTMLInputElement, HTMLSelectElement, document */

import { byId, inputValue, setStatus } from "../services/ui";
import {
  DEFAULT_RECON_SETTINGS,
  ReconSettings,
  loadReconSettings,
  saveReconSettings,
} from "../services/settings";
import {
  DEFAULT_THRESHOLDS,
  DEFAULT_TOLERANCES,
  MatchingMode,
  ReconciliationResult,
  Tolerances,
  reconcile,
} from "../services/reconciliation";
import {
  ColumnMapping,
  RangeDescriptor,
  autoDetectColumns,
  captureCurrentSelection,
  looksLikeHeader,
  normalizeRows,
  readRangeValues,
  writeReconciliationSheet,
} from "../services/reconcileExcel";

type Side = "a" | "b";

interface SideState {
  descriptor: RangeDescriptor | null;
  mapping: ColumnMapping | null;
}

const state: { [K in Side]: SideState } = {
  a: { descriptor: null, mapping: null },
  b: { descriptor: null, mapping: null },
};

export function initReconciliationTab(): void {
  // Range capture
  byId<HTMLButtonElement>("recon-capture-a").addEventListener("click", () => void captureFor("a"));
  byId<HTMLButtonElement>("recon-capture-b").addEventListener("click", () => void captureFor("b"));

  // First-row-is-header toggles re-run auto-detection
  byId<HTMLInputElement>("recon-header-a").addEventListener("change", () => recomputeMapping("a"));
  byId<HTMLInputElement>("recon-header-b").addEventListener("change", () => recomputeMapping("b"));

  // Matching mode presets
  byId<HTMLSelectElement>("recon-mode").addEventListener("change", onModeChange);

  // Any manual tolerance edit switches the mode select to "custom" so the user
  // knows they're no longer on a preset.
  const tolInputIds = ["recon-amt-fixed", "recon-amt-pct", "recon-date-days"];
  for (let i = 0; i < tolInputIds.length; i++) {
    byId<HTMLInputElement>(tolInputIds[i]).addEventListener("input", () => {
      const sel = byId<HTMLSelectElement>("recon-mode");
      if (sel.value !== "custom") sel.value = "custom";
    });
  }

  // Actions
  byId<HTMLButtonElement>("recon-run-btn").addEventListener(
    "click",
    () => void runReconciliation()
  );
  byId<HTMLButtonElement>("recon-save-defaults").addEventListener(
    "click",
    () => void saveDefaults()
  );

  // Hydrate form from persisted defaults (or baked-in defaults if none saved).
  applyReconSettings(loadReconSettings());
}

/* --------------------- Tolerance inputs & presets --------------------- */

function onModeChange(): void {
  const mode = byId<HTMLSelectElement>("recon-mode").value as MatchingMode;
  if (mode === "custom") return; // leave the user's values alone
  const tol = DEFAULT_TOLERANCES[mode];
  byId<HTMLInputElement>("recon-amt-fixed").value = String(tol.amountFixed);
  byId<HTMLInputElement>("recon-amt-pct").value = String(tol.amountPct);
  byId<HTMLInputElement>("recon-date-days").value = String(tol.dateDays);
}

function readFormTolerances(): { mode: MatchingMode; tol: Tolerances } {
  const mode = byId<HTMLSelectElement>("recon-mode").value as MatchingMode;
  const tol: Tolerances = {
    amountFixed: parseNonNegative("recon-amt-fixed", DEFAULT_RECON_SETTINGS.amountFixed),
    amountPct: parseNonNegative("recon-amt-pct", DEFAULT_RECON_SETTINGS.amountPct),
    dateDays: Math.floor(parseNonNegative("recon-date-days", DEFAULT_RECON_SETTINGS.dateDays)),
  };
  return { mode, tol };
}

function parseNonNegative(id: string, fallback: number): number {
  const raw = inputValue(id);
  const n = parseFloat(raw);
  return isFinite(n) && n >= 0 ? n : fallback;
}

function applyReconSettings(s: ReconSettings): void {
  byId<HTMLSelectElement>("recon-mode").value = s.mode;
  byId<HTMLInputElement>("recon-amt-fixed").value = String(s.amountFixed);
  byId<HTMLInputElement>("recon-amt-pct").value = String(s.amountPct);
  byId<HTMLInputElement>("recon-date-days").value = String(s.dateDays);
}

/* ------------------- Range capture & column mapping UI ------------------- */

async function captureFor(side: Side): Promise<void> {
  const btn = byId<HTMLButtonElement>(`recon-capture-${side}`);
  btn.disabled = true;
  setStatus(`Capturing Source ${side.toUpperCase()} selection…`, "loading");
  try {
    const desc = await captureCurrentSelection();
    state[side].descriptor = desc;

    const hasHeader = looksLikeHeader(desc.firstRow);
    byId<HTMLInputElement>(`recon-header-${side}`).checked = hasHeader;

    state[side].mapping = autoDetectColumns(desc.firstRow, desc.sampleRow, hasHeader);

    renderRangeSummary(side);
    populateColumnDropdowns(side);
    setStatus(`Source ${side.toUpperCase()} captured: ${desc.address}`, "success");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
  }
}

function renderRangeSummary(side: Side): void {
  const desc = state[side].descriptor;
  const el = byId(`recon-range-${side}`);
  if (!desc) {
    el.textContent = "No range selected.";
    el.classList.remove("is-set");
    byId(`recon-map-${side}`).hidden = true;
    return;
  }
  el.textContent = `${desc.address}  ·  ${desc.rowCount} rows × ${desc.columnCount} cols`;
  el.classList.add("is-set");
  byId(`recon-map-${side}`).hidden = false;
}

function populateColumnDropdowns(side: Side): void {
  const desc = state[side].descriptor;
  const mapping = state[side].mapping;
  if (!desc || !mapping) return;
  const labels = buildColumnLabels(desc, mapping.hasHeader);
  fillSelect(`recon-col-${side}-date`, labels, mapping.dateCol, (v) => {
    if (state[side].mapping) state[side].mapping!.dateCol = v;
  });
  fillSelect(`recon-col-${side}-amount`, labels, mapping.amountCol, (v) => {
    if (state[side].mapping) state[side].mapping!.amountCol = v;
  });
  fillSelect(`recon-col-${side}-desc`, labels, mapping.descCol, (v) => {
    if (state[side].mapping) state[side].mapping!.descCol = v;
  });
}

function buildColumnLabels(desc: RangeDescriptor, hasHeader: boolean): string[] {
  const labels: string[] = new Array(desc.columnCount);
  for (let i = 0; i < desc.columnCount; i++) {
    const letter = columnIndexToLetter(i);
    const headerCell = desc.firstRow[i];
    if (hasHeader && typeof headerCell === "string" && headerCell.trim()) {
      const short = headerCell.trim().slice(0, 30);
      labels[i] = `${letter} — ${short}`;
    } else {
      labels[i] = `Column ${letter}`;
    }
  }
  return labels;
}

function columnIndexToLetter(i: number): string {
  let s = "";
  let n = i;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

function fillSelect(
  id: string,
  options: string[],
  selected: number,
  onChange: (v: number) => void
): void {
  const sel = byId<HTMLSelectElement>(id);
  sel.innerHTML = "";
  for (let i = 0; i < options.length; i++) {
    const opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = options[i];
    if (i === selected) opt.selected = true;
    sel.appendChild(opt);
  }
  // onchange (not addEventListener) so repeated populates don't stack handlers.
  sel.onchange = () => onChange(parseInt(sel.value, 10));
}

function recomputeMapping(side: Side): void {
  const desc = state[side].descriptor;
  if (!desc) return;
  const hasHeader = byId<HTMLInputElement>(`recon-header-${side}`).checked;
  state[side].mapping = autoDetectColumns(desc.firstRow, desc.sampleRow, hasHeader);
  populateColumnDropdowns(side);
}

/* -------------------------- Run reconciliation -------------------------- */

async function runReconciliation(): Promise<void> {
  const a = state.a;
  const b = state.b;
  if (!a.descriptor || !a.mapping) {
    setStatus("Capture a range for Source A first.", "error");
    return;
  }
  if (!b.descriptor || !b.mapping) {
    setStatus("Capture a range for Source B first.", "error");
    return;
  }

  const { mode, tol } = readFormTolerances();
  // For "custom" mode we use the Normal thresholds so users still get the
  // three-tier Matched / Possible / Unmatched output without configuring it.
  const thresholds = mode === "custom" ? DEFAULT_THRESHOLDS.normal : DEFAULT_THRESHOLDS[mode];

  const runBtn = byId<HTMLButtonElement>("recon-run-btn");
  runBtn.disabled = true;
  setStatus("Reconciling…", "loading");
  try {
    const [valuesA, valuesB] = await Promise.all([
      readRangeValues(a.descriptor),
      readRangeValues(b.descriptor),
    ]);
    const rowsA = normalizeRows(valuesA, a.mapping, parseFirstExcelRow(a.descriptor.a1));
    const rowsB = normalizeRows(valuesB, b.mapping, parseFirstExcelRow(b.descriptor.a1));

    if (!rowsA.length || !rowsB.length) {
      throw new Error(
        "One of the ranges has no data rows after normalization. Check your column mapping and the 'First row is a header' toggle."
      );
    }

    const result = reconcile(rowsA, rowsB, tol, thresholds);
    const sheetName = await writeReconciliationSheet(result, {
      mode,
      tolerances: tol,
      descA: a.descriptor,
      descB: b.descriptor,
    });

    renderSummaryCard(result);
    setStatus(
      `Done. Matched ${result.summary.matched} · Possible ${result.summary.possible} · Only in A ${result.summary.unmatchedA} · Only in B ${result.summary.unmatchedB}. Wrote "${sheetName}".`,
      "success"
    );
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    runBtn.disabled = false;
  }
}

/** Extract the 1-based row number from an A1-style range reference. */
function parseFirstExcelRow(a1: string): number {
  const m = /(\d+)/.exec(a1);
  return m ? parseInt(m[1], 10) : 1;
}

function renderSummaryCard(result: ReconciliationResult): void {
  const el = byId("recon-summary");
  const { matched, possible, unmatchedA, unmatchedB, totalA, totalB } = result.summary;
  el.innerHTML = `
    <h4>Last run</h4>
    <div class="recon__stat-grid">
      <div class="recon__stat recon__stat--matched">
        <div class="recon__stat-value">${matched}</div>
        <div class="recon__stat-label">Matched</div>
      </div>
      <div class="recon__stat recon__stat--possible">
        <div class="recon__stat-value">${possible}</div>
        <div class="recon__stat-label">Possible</div>
      </div>
      <div class="recon__stat recon__stat--unmatched">
        <div class="recon__stat-value">${unmatchedA}</div>
        <div class="recon__stat-label">Only in A</div>
      </div>
      <div class="recon__stat recon__stat--unmatched">
        <div class="recon__stat-value">${unmatchedB}</div>
        <div class="recon__stat-label">Only in B</div>
      </div>
    </div>
    <p class="recon__summary-foot">Totals: A = ${totalA} · B = ${totalB}</p>
  `;
  el.hidden = false;
}

/* ------------------------------ Defaults ------------------------------ */

async function saveDefaults(): Promise<void> {
  const { mode, tol } = readFormTolerances();
  try {
    await saveReconSettings({
      mode,
      amountFixed: tol.amountFixed,
      amountPct: tol.amountPct,
      dateDays: tol.dateDays,
    });
    setStatus("Reconciliation defaults saved for this workbook.", "success");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  }
}
