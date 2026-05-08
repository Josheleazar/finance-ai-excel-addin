/*
 * Reconciliation tab — orchestrates the end-to-end user flow.
 *
 * Flow:
 *   1. User captures two Excel ranges (Source A & B) via "Use selection".
 *   2. User defines 1..6 matching fields. Each field picks one column from A
 *      and one from B, plus a type (exact / numeric / date / fuzzy), weight,
 *      and an optional Required flag. Required fields hard-gate non-matches
 *      by default; the advanced panel offers a "downgrade" opt-in instead.
 *   3. User picks a Sensitivity (Strict / Normal / Loose) which controls the
 *      match / possible confidence thresholds.
 *   4. "Run Reconciliation" reads both ranges, scores every candidate pair,
 *      greedy-assigns matches, and writes a results worksheet.
 *   5. "Save as Default" persists the current field configuration.
 *
 * The matching engine lives in services/reconciliation (pure, testable).
 * Excel I/O lives in services/reconcileExcel. This file is the DOM glue.
 */

/* global HTMLButtonElement, HTMLInputElement, HTMLSelectElement, HTMLElement, document */

import { byId, setStatus } from "../services/ui";
import { PersistedField, loadReconSettings, saveReconSettings } from "../services/settings";
import {
  FieldConfig,
  FieldType,
  FieldWeight,
  MAX_FIELDS,
  ReconciliationResult,
  SENSITIVITY_THRESHOLDS,
  Sensitivity,
  defaultToleranceForType,
  reconcile,
} from "../services/reconciliation";
import {
  RangeDescriptor,
  buildRawRows,
  captureCurrentSelection,
  columnIndexToLetter,
  detectColumnForField,
  looksLikeHeader,
  readRangeValues,
  writeReconciliationSheet,
} from "../services/reconcileExcel";

type Side = "a" | "b";

const state: {
  a: { descriptor: RangeDescriptor | null };
  b: { descriptor: RangeDescriptor | null };
  fields: FieldConfig[];
  sensitivity: Sensitivity;
  /** Which fields have their "Advanced" panel expanded. UI-only. */
  expanded: { [fieldId: string]: boolean };
} = {
  a: { descriptor: null },
  b: { descriptor: null },
  fields: [],
  sensitivity: "normal",
  expanded: {},
};

export function initReconciliationTab(): void {
  // Range capture
  byId<HTMLButtonElement>("recon-capture-a").addEventListener("click", () => void captureFor("a"));
  byId<HTMLButtonElement>("recon-capture-b").addEventListener("click", () => void captureFor("b"));

  // Header toggles — re-run auto-detection for any field whose column is unset.
  byId<HTMLInputElement>("recon-header-a").addEventListener("change", () => {
    autoDetectMissingColumns("a");
    renderFields();
  });
  byId<HTMLInputElement>("recon-header-b").addEventListener("change", () => {
    autoDetectMissingColumns("b");
    renderFields();
  });

  // Sensitivity
  byId<HTMLSelectElement>("recon-sensitivity").addEventListener("change", () => {
    state.sensitivity = byId<HTMLSelectElement>("recon-sensitivity").value as Sensitivity;
  });

  // Field list actions
  byId<HTMLButtonElement>("recon-add-field").addEventListener("click", () => addField());

  // Run / Save
  byId<HTMLButtonElement>("recon-run-btn").addEventListener(
    "click",
    () => void runReconciliation()
  );
  byId<HTMLButtonElement>("recon-save-defaults").addEventListener(
    "click",
    () => void saveDefaults()
  );

  // Hydrate from persisted settings (or baked-in defaults if none saved).
  const saved = loadReconSettings();
  state.sensitivity = saved.sensitivity;
  state.fields = saved.fields.map(persistedToConfig);
  byId<HTMLSelectElement>("recon-sensitivity").value = state.sensitivity;
  renderFields();
}

/* ---------------------- Field model helpers ---------------------- */

function makeFieldId(): string {
  return `f${Date.now().toString(36)}${Math.floor(Math.random() * 1e6).toString(36)}`;
}

function persistedToConfig(p: PersistedField): FieldConfig {
  return {
    id: p.id,
    label: p.label,
    type: p.type,
    colA: -1,
    colB: -1,
    weight: p.weight,
    required: p.required,
    downgradeOnFail: p.downgradeOnFail,
    tolerance: { ...p.tolerance },
  };
}

function configToPersisted(f: FieldConfig): PersistedField {
  return {
    id: f.id,
    label: f.label,
    type: f.type,
    weight: f.weight,
    required: f.required,
    downgradeOnFail: f.downgradeOnFail,
    tolerance: { ...f.tolerance },
  };
}

function addField(): void {
  if (state.fields.length >= MAX_FIELDS) return;
  const n = state.fields.length + 1;
  const field: FieldConfig = {
    id: makeFieldId(),
    label: `Field ${n}`,
    type: "fuzzy",
    colA: -1,
    colB: -1,
    weight: "medium",
    required: false,
    downgradeOnFail: false,
    tolerance: defaultToleranceForType("fuzzy"),
  };
  state.fields.push(field);
  autoDetectMissingColumns("a");
  autoDetectMissingColumns("b");
  renderFields();
}

function removeField(fieldId: string): void {
  state.fields = state.fields.filter((f) => f.id !== fieldId);
  delete state.expanded[fieldId];
  renderFields();
}

/** Fill in colA/colB for any field whose current pick is unset or out-of-range.
 *  Tracks columns already claimed by other fields on the same side so the
 *  default seed of 3 fields doesn't collapse onto a single column when the
 *  heuristics can't decide. */
function autoDetectMissingColumns(side: Side): void {
  const desc = state[side].descriptor;
  if (!desc) return;
  const hasHeader = byId<HTMLInputElement>(`recon-header-${side}`).checked;

  // Seed `taken` with columns already manually/previously assigned on this side.
  const taken: { [col: number]: boolean } = {};
  for (let k = 0; k < state.fields.length; k++) {
    const current = side === "a" ? state.fields[k].colA : state.fields[k].colB;
    if (current >= 0 && current < desc.columnCount) taken[current] = true;
  }

  for (let k = 0; k < state.fields.length; k++) {
    const f = state.fields[k];
    const current = side === "a" ? f.colA : f.colB;
    if (current < 0 || current >= desc.columnCount) {
      const suggested = detectColumnForField(f, desc, hasHeader, taken);
      if (side === "a") f.colA = suggested;
      else f.colB = suggested;
      if (suggested >= 0) taken[suggested] = true;
    }
  }
}

/* -------------------------- Range capture -------------------------- */

async function captureFor(side: Side): Promise<void> {
  const btn = byId<HTMLButtonElement>(`recon-capture-${side}`);
  btn.disabled = true;
  setStatus(`Capturing Source ${side.toUpperCase()} selection…`, "loading");
  try {
    const desc = await captureCurrentSelection();
    state[side].descriptor = desc;

    // Best-effort header detection as a starting point; user can toggle after.
    const hasHeader = looksLikeHeader(desc.firstRow);
    byId<HTMLInputElement>(`recon-header-${side}`).checked = hasHeader;

    // Any existing column pick that no longer fits is reset to -1, then
    // auto-detected so a sensible default appears in the dropdown.
    for (let k = 0; k < state.fields.length; k++) {
      const f = state.fields[k];
      if (side === "a" && f.colA >= desc.columnCount) f.colA = -1;
      if (side === "b" && f.colB >= desc.columnCount) f.colB = -1;
    }
    autoDetectMissingColumns(side);

    renderRangeSummary(side);
    renderFields();
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
    return;
  }
  el.textContent = `${desc.address}  ·  ${desc.rowCount} rows × ${desc.columnCount} cols`;
  el.classList.add("is-set");
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

/* --------------------------- Render fields --------------------------- */

function renderFields(): void {
  const container = byId("recon-fields");
  container.innerHTML = "";
  for (let k = 0; k < state.fields.length; k++) {
    container.appendChild(renderFieldRow(state.fields[k]));
  }
  byId("recon-fields-count").textContent = `${state.fields.length} of ${MAX_FIELDS}`;
  byId<HTMLButtonElement>("recon-add-field").disabled = state.fields.length >= MAX_FIELDS;

  // Show an empty-state hint when there are no fields at all.
  const empty = byId("recon-fields-empty");
  if (empty) empty.hidden = state.fields.length > 0;
}

function renderFieldRow(f: FieldConfig): HTMLElement {
  const row = document.createElement("div");
  row.className = "recon-field";
  row.setAttribute("data-field-id", f.id);

  /* Head: label input + type select + remove button */
  const head = document.createElement("div");
  head.className = "recon-field__head";

  const labelInput = document.createElement("input");
  labelInput.type = "text";
  labelInput.className = "recon-field__label field__input";
  labelInput.value = f.label;
  labelInput.placeholder = "Field name";
  labelInput.addEventListener("input", () => {
    f.label = labelInput.value;
  });
  head.appendChild(labelInput);

  const typeSelect = document.createElement("select");
  typeSelect.className = "recon-field__type field__input";
  const typeOptions: Array<[FieldType, string]> = [
    ["exact", "Exact"],
    ["numeric", "Numeric"],
    ["date", "Date"],
    ["fuzzy", "Fuzzy text"],
  ];
  for (let i = 0; i < typeOptions.length; i++) {
    const opt = document.createElement("option");
    opt.value = typeOptions[i][0];
    opt.textContent = typeOptions[i][1];
    if (typeOptions[i][0] === f.type) opt.selected = true;
    typeSelect.appendChild(opt);
  }
  typeSelect.addEventListener("change", () => {
    f.type = typeSelect.value as FieldType;
    // Replace tolerance with the new type's defaults — keeping old keys would
    // be misleading.
    f.tolerance = defaultToleranceForType(f.type);
    renderFields();
  });
  head.appendChild(typeSelect);

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "recon-field__remove icon-btn";
  removeBtn.title = "Remove field";
  removeBtn.setAttribute("aria-label", "Remove field");
  removeBtn.textContent = "×";
  removeBtn.addEventListener("click", () => removeField(f.id));
  head.appendChild(removeBtn);

  row.appendChild(head);

  /* Body: column pickers + meta + optional advanced panel */
  const body = document.createElement("div");
  body.className = "recon-field__body";

  const colsRow = document.createElement("div");
  colsRow.className = "row row--2 recon-field__cols";
  colsRow.appendChild(renderColumnPicker(f, "a"));
  colsRow.appendChild(renderColumnPicker(f, "b"));
  body.appendChild(colsRow);

  const meta = document.createElement("div");
  meta.className = "recon-field__meta";
  meta.appendChild(renderWeightChips(f));

  const reqLabel = document.createElement("label");
  reqLabel.className = "recon-field__toggle";
  const reqInput = document.createElement("input");
  reqInput.type = "checkbox";
  reqInput.checked = f.required;
  reqInput.addEventListener("change", () => {
    f.required = reqInput.checked;
    // Re-render so the advanced panel's "On failure" block appears/disappears.
    renderFields();
  });
  reqLabel.appendChild(reqInput);
  reqLabel.appendChild(document.createTextNode(" Required"));
  meta.appendChild(reqLabel);

  const advToggle = document.createElement("button");
  advToggle.type = "button";
  advToggle.className = "recon-field__adv-toggle";
  const isExpanded = !!state.expanded[f.id];
  advToggle.textContent = (isExpanded ? "▾" : "▸") + " Advanced";
  advToggle.addEventListener("click", () => {
    state.expanded[f.id] = !state.expanded[f.id];
    renderFields();
  });
  meta.appendChild(advToggle);

  body.appendChild(meta);

  if (isExpanded) {
    body.appendChild(renderAdvancedPanel(f));
  }

  row.appendChild(body);
  return row;
}

function renderColumnPicker(f: FieldConfig, side: Side): HTMLElement {
  const desc = state[side].descriptor;
  const wrap = document.createElement("label");
  wrap.className = "field";

  const span = document.createElement("span");
  span.className = "field__label";
  span.textContent = `Col ${side.toUpperCase()}`;
  wrap.appendChild(span);

  const select = document.createElement("select");
  select.className = "field__input";

  if (!desc) {
    const opt = document.createElement("option");
    opt.value = "-1";
    opt.textContent = `— capture Source ${side.toUpperCase()} —`;
    opt.disabled = true;
    opt.selected = true;
    select.appendChild(opt);
    select.disabled = true;
  } else {
    const hasHeader = byId<HTMLInputElement>(`recon-header-${side}`).checked;
    const labels = buildColumnLabels(desc, hasHeader);
    const unset = document.createElement("option");
    unset.value = "-1";
    unset.textContent = "— unset —";
    select.appendChild(unset);
    for (let i = 0; i < labels.length; i++) {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = labels[i];
      select.appendChild(opt);
    }
    const current = side === "a" ? f.colA : f.colB;
    select.value = String(current >= 0 && current < labels.length ? current : -1);
    select.addEventListener("change", () => {
      const v = parseInt(select.value, 10);
      if (side === "a") f.colA = v;
      else f.colB = v;
    });
  }
  wrap.appendChild(select);
  return wrap;
}

function renderWeightChips(f: FieldConfig): HTMLElement {
  const wrap = document.createElement("div");
  wrap.className = "recon-field__weight";

  const label = document.createElement("span");
  label.className = "recon-field__weight-label";
  label.textContent = "Weight";
  wrap.appendChild(label);

  const group = document.createElement("div");
  group.className = "recon-field__weight-chips";
  const weights: Array<[FieldWeight, string]> = [
    ["low", "Low"],
    ["medium", "Med"],
    ["high", "High"],
  ];
  for (let i = 0; i < weights.length; i++) {
    const [w, l] = weights[i];
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chip" + (f.weight === w ? " is-active" : "");
    btn.textContent = l;
    btn.addEventListener("click", () => {
      f.weight = w;
      renderFields();
    });
    group.appendChild(btn);
  }
  wrap.appendChild(group);
  return wrap;
}

function renderAdvancedPanel(f: FieldConfig): HTMLElement {
  const panel = document.createElement("div");
  panel.className = "recon-field__advanced";

  /* Type-specific tolerance inputs */
  const tolRow = document.createElement("div");
  tolRow.className = "recon-field__tol-row";

  if (f.type === "numeric") {
    tolRow.appendChild(
      renderNumberInput("± Absolute", f.tolerance.amountFixed || 0, 0.01, (v) => {
        f.tolerance.amountFixed = v;
      })
    );
    tolRow.appendChild(
      renderNumberInput("± Percent %", f.tolerance.amountPct || 0, 0.1, (v) => {
        f.tolerance.amountPct = v;
      })
    );
  } else if (f.type === "date") {
    tolRow.appendChild(
      renderNumberInput("± Days", f.tolerance.dateDays || 0, 1, (v) => {
        f.tolerance.dateDays = Math.floor(v);
      })
    );
  } else if (f.type === "fuzzy") {
    const current = f.tolerance.minSimilarity === undefined ? 0.6 : f.tolerance.minSimilarity;
    tolRow.appendChild(
      renderNumberInput("Min similarity (0–1)", current, 0.05, (v) => {
        f.tolerance.minSimilarity = Math.max(0, Math.min(1, v));
      })
    );
  } else {
    const toggle = document.createElement("label");
    toggle.className = "recon-field__toggle";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!f.tolerance.caseSensitive;
    cb.addEventListener("change", () => {
      f.tolerance.caseSensitive = cb.checked;
    });
    toggle.appendChild(cb);
    toggle.appendChild(document.createTextNode(" Case-sensitive"));
    tolRow.appendChild(toggle);
  }
  panel.appendChild(tolRow);

  /* Required-failure behavior radios — only shown when Required is on. */
  if (f.required) {
    const failWrap = document.createElement("fieldset");
    failWrap.className = "recon-field__fail";
    const legend = document.createElement("legend");
    legend.textContent = "If this field fails";
    failWrap.appendChild(legend);

    const name = `fail-${f.id}`;
    failWrap.appendChild(
      renderRadio(name, "Block match entirely (hard gate)", !f.downgradeOnFail, () => {
        f.downgradeOnFail = false;
      })
    );
    failWrap.appendChild(
      renderRadio(name, "Downgrade to Possible Match", f.downgradeOnFail, () => {
        f.downgradeOnFail = true;
      })
    );
    panel.appendChild(failWrap);
  }

  return panel;
}

function renderNumberInput(
  label: string,
  value: number,
  step: number,
  onChange: (v: number) => void
): HTMLElement {
  const wrap = document.createElement("label");
  wrap.className = "field";
  const span = document.createElement("span");
  span.className = "field__label";
  span.textContent = label;
  wrap.appendChild(span);
  const input = document.createElement("input");
  input.type = "number";
  input.className = "field__input";
  input.min = "0";
  input.step = String(step);
  input.value = String(value);
  input.addEventListener("input", () => {
    const n = parseFloat(input.value);
    if (isFinite(n) && n >= 0) onChange(n);
  });
  wrap.appendChild(input);
  return wrap;
}

function renderRadio(
  name: string,
  label: string,
  checked: boolean,
  onChange: () => void
): HTMLElement {
  const wrap = document.createElement("label");
  wrap.className = "recon-field__radio";
  const radio = document.createElement("input");
  radio.type = "radio";
  radio.name = name;
  radio.checked = checked;
  radio.addEventListener("change", () => {
    if (radio.checked) onChange();
  });
  wrap.appendChild(radio);
  wrap.appendChild(document.createTextNode(" " + label));
  return wrap;
}

/* -------------------------- Run reconciliation -------------------------- */

async function runReconciliation(): Promise<void> {
  if (!state.a.descriptor) {
    setStatus("Capture a range for Source A first.", "error");
    return;
  }
  if (!state.b.descriptor) {
    setStatus("Capture a range for Source B first.", "error");
    return;
  }
  if (!state.fields.length) {
    setStatus("Add at least one matching field.", "error");
    return;
  }

  const incomplete = state.fields.filter((f) => f.colA < 0 || f.colB < 0);
  if (incomplete.length) {
    const names = incomplete.map((f) => f.label || "(unnamed)").join(", ");
    setStatus(`These fields still need a column on both sides: ${names}.`, "error");
    return;
  }

  const runBtn = byId<HTMLButtonElement>("recon-run-btn");
  runBtn.disabled = true;
  setStatus("Reconciling…", "loading");
  try {
    const hasHeaderA = byId<HTMLInputElement>("recon-header-a").checked;
    const hasHeaderB = byId<HTMLInputElement>("recon-header-b").checked;
    const [valuesA, valuesB] = await Promise.all([
      readRangeValues(state.a.descriptor),
      readRangeValues(state.b.descriptor),
    ]);
    const rowsA = buildRawRows(valuesA, hasHeaderA, parseFirstExcelRow(state.a.descriptor.a1));
    const rowsB = buildRawRows(valuesB, hasHeaderB, parseFirstExcelRow(state.b.descriptor.a1));

    if (!rowsA.length || !rowsB.length) {
      throw new Error(
        "One of the ranges has no data rows after normalization. Check the 'First row is a header' toggle."
      );
    }

    const thresholds = SENSITIVITY_THRESHOLDS[state.sensitivity];
    const result = reconcile(rowsA, rowsB, state.fields, thresholds);
    const sheetName = await writeReconciliationSheet(result, {
      sensitivity: state.sensitivity,
      fields: state.fields,
      descA: state.a.descriptor,
      descB: state.b.descriptor,
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
  try {
    await saveReconSettings({
      sensitivity: state.sensitivity,
      fields: state.fields.map(configToPersisted),
    });
    setStatus("Reconciliation defaults saved for this workbook.", "success");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  }
}
