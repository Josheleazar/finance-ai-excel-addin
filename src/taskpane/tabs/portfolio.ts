/*
 * Portfolio Tracker tab.
 *
 * Flow:
 *   1. On mount, load the Portfolio worksheet (if any) into in-memory state.
 *   2. User adds rows manually — typing a symbol triggers a Finnhub lookup to
 *      auto-fill Company Name, Sector, and Current Price. Quantity and cost
 *      basis are required.
 *   3. Alternative entry: "Import from Broker CSV" opens a file picker,
 *      parses the CSV, and walks the user through a smart-suggested column
 *      mapping before appending / merging rows into state.
 *   4. Every state mutation is persisted by re-writing the entire Portfolio
 *      worksheet — cheap because the dataset is small (typically < 200 rows)
 *      and keeps the sheet as the single source of truth.
 *   5. "Refresh All Prices" pulls a fresh quote for each symbol in parallel.
 *   6. "Reconcile Portfolio" captures the user's current Excel selection as
 *      a broker export and runs a purpose-built comparison (Symbol + Qty +
 *      Cost Basis), writing results to a new worksheet.
 */

/* global HTMLButtonElement, HTMLInputElement, HTMLSelectElement, HTMLElement, KeyboardEvent, document, window, FileReader, File */

import { getCompanyProfile, getQuote } from "../services/finnhub";
import { loadApiKey } from "../services/settings";
import { byId, setStatus } from "../services/ui";
import {
  ColumnMapping,
  HOLDING_FIELDS,
  HOLDING_FIELD_LABELS,
  Holding,
  HoldingField,
  REQUIRED_IMPORT_FIELDS,
  applyMapping,
  computeHoldingMetrics,
  computeTotals,
  mergeHoldings,
  normalizeSymbol,
  parseCsv,
  parseLooseNumber,
  suggestColumnMapping,
} from "../services/portfolio";
import {
  BrokerRange,
  captureBrokerSelection,
  readPortfolioFromSheet,
  reconcilePortfolio,
  summarizePortfolioRecon,
  writePortfolioReconSheet,
  writePortfolioToSheet,
} from "../services/portfolioExcel";

interface PendingCsv {
  fileName: string;
  headers: string[];
  rows: string[][];
  mapping: ColumnMapping;
}

interface PendingBroker {
  range: BrokerRange;
  mapping: { symbolCol: number; quantityCol: number; costCol: number };
}

/** Cached result of the most recent "lookup on blur" for the Add form, keyed
 *  by symbol. Consumed once by handleAddHolding and cleared on successful add. */
interface PendingLookup {
  symbol: string;
  name: string;
  sector: string;
  price: number | null;
}

const state: {
  holdings: Holding[];
  pendingCsv: PendingCsv | null;
  pendingBroker: PendingBroker | null;
  pendingLookup: PendingLookup | null;
  /** When set (ms timestamp), the next Clear click executes instead of prompting. */
  clearArmedUntil: number;
} = {
  holdings: [],
  pendingCsv: null,
  pendingBroker: null,
  pendingLookup: null,
  clearArmedUntil: 0,
};

export function initPortfolioTab(): void {
  // Manual entry
  byId<HTMLButtonElement>("pf-add-btn").addEventListener("click", () => void handleAddHolding());
  byId<HTMLInputElement>("pf-add-symbol").addEventListener("blur", () => void handleSymbolAutoFill());
  byId<HTMLInputElement>("pf-add-symbol").addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      void handleSymbolAutoFill();
    }
  });

  // CSV import
  byId<HTMLButtonElement>("pf-csv-btn").addEventListener("click", () => {
    byId<HTMLInputElement>("pf-csv-file").click();
  });
  byId<HTMLInputElement>("pf-csv-file").addEventListener("change", (e) => {
    const input = e.target as HTMLInputElement;
    if (input.files && input.files.length) {
      void handleCsvFile(input.files[0]);
      input.value = ""; // allow re-selecting the same file
    }
  });
  byId<HTMLButtonElement>("pf-csv-cancel").addEventListener("click", () => {
    state.pendingCsv = null;
    renderCsvMapper();
  });
  byId<HTMLButtonElement>("pf-csv-confirm").addEventListener("click", () => void confirmCsvImport());

  // Toolbar actions
  byId<HTMLButtonElement>("pf-refresh-btn").addEventListener("click", () => void refreshAllPrices());
  byId<HTMLButtonElement>("pf-recon-btn").addEventListener("click", () => void startReconciliation());
  byId<HTMLButtonElement>("pf-recon-run").addEventListener("click", () => void runReconciliation());
  byId<HTMLButtonElement>("pf-recon-cancel").addEventListener("click", () => {
    state.pendingBroker = null;
    renderReconPanel();
  });
  byId<HTMLButtonElement>("pf-clear-btn").addEventListener("click", () => void clearAllHoldings());

  // Initial load from sheet (async, non-blocking).
  void bootstrap();
}

async function bootstrap(): Promise<void> {
  try {
    const holdings = await readPortfolioFromSheet();
    state.holdings = holdings;
    renderAll();
  } catch (err) {
    // Non-fatal: user may not have opened a workbook yet, or sheet is missing.
    // eslint-disable-next-line no-console
    console.warn("Portfolio bootstrap: could not read existing sheet", err);
    renderAll();
  }
}

/* ---------------------------- Manual entry ---------------------------- */

async function handleSymbolAutoFill(): Promise<void> {
  const input = byId<HTMLInputElement>("pf-add-symbol");
  const symbol = normalizeSymbol(input.value);
  if (!symbol) return;
  input.value = symbol;
  const preview = byId("pf-add-preview");
  preview.textContent = "Looking up…";
  preview.classList.remove("is-error");

  const apiKey = loadApiKey();
  if (!apiKey) {
    preview.textContent = "Set your Finnhub API key in Settings to auto-fill.";
    preview.classList.add("is-error");
    return;
  }

  try {
    // Profile + quote in parallel. Either one failing shouldn't block the other.
    const [profileRes, quoteRes] = await Promise.allSettled([
      getCompanyProfile(symbol, apiKey),
      getQuote(symbol, apiKey),
    ]);
    const parts: string[] = [symbol];
    let name = "";
    let sector = "";
    let price: number | undefined;
    if (profileRes.status === "fulfilled") {
      name = profileRes.value.name || "";
      sector = profileRes.value.finnhubIndustry || "";
      if (name) parts.push(name);
      if (sector) parts.push(sector);
    }
    if (quoteRes.status === "fulfilled") {
      price = quoteRes.value.c;
      parts.push(`$${price.toFixed(2)}`);
    }
    preview.textContent = parts.join(" · ");
    // Cache in module state so handleAddHolding can pick it up without re-fetching.
    state.pendingLookup = {
      symbol,
      name,
      sector,
      price: price === undefined ? null : price,
    };
  } catch (err) {
    preview.textContent = err instanceof Error ? err.message : String(err);
    preview.classList.add("is-error");
    state.pendingLookup = null;
  }
}

async function handleAddHolding(): Promise<void> {
  const symbol = normalizeSymbol(byId<HTMLInputElement>("pf-add-symbol").value);
  if (!symbol) {
    setStatus("Enter a ticker symbol.", "error");
    return;
  }
  const qty = parseLooseNumber(byId<HTMLInputElement>("pf-add-qty").value);
  if (qty === null || qty <= 0) {
    setStatus("Quantity must be a positive number.", "error");
    return;
  }
  const cost = parseLooseNumber(byId<HTMLInputElement>("pf-add-cost").value);
  if (cost === null || cost < 0) {
    setStatus("Avg cost basis must be a non-negative number.", "error");
    return;
  }

  // Only trust the cached lookup if it's for this exact symbol.
  const cached =
    state.pendingLookup && state.pendingLookup.symbol === symbol
      ? state.pendingLookup
      : null;

  const holding: Holding = {
    symbol,
    quantity: qty,
    avgCostBasis: cost,
  };
  if (cached) {
    if (cached.name) holding.companyName = cached.name;
    if (cached.sector) holding.sector = cached.sector;
    if (cached.price !== null) {
      holding.currentPrice = cached.price;
      holding.lastRefreshedAt = Date.now();
    }
  }

  state.holdings = mergeHoldings(state.holdings, [holding]);

  // If the cached preview was missing or incomplete, fetch in the background
  // so the row enriches without blocking the add.
  if (!cached || !cached.name || cached.price === null) {
    void enrichHolding(symbol);
  }

  // Reset inputs
  byId<HTMLInputElement>("pf-add-symbol").value = "";
  byId<HTMLInputElement>("pf-add-qty").value = "";
  byId<HTMLInputElement>("pf-add-cost").value = "";
  byId("pf-add-preview").textContent = "";
  state.pendingLookup = null;

  await persistAndRender();
  setStatus(`Added ${symbol}.`, "success");
}

/** Fetch profile + quote for a symbol and patch the in-memory holding.
 *  Used when the user adds a row without waiting for the blur/Enter lookup. */
async function enrichHolding(symbol: string): Promise<void> {
  const apiKey = loadApiKey();
  if (!apiKey) return;
  try {
    const [profileRes, quoteRes] = await Promise.allSettled([
      getCompanyProfile(symbol, apiKey),
      getQuote(symbol, apiKey),
    ]);
    const idx = state.holdings.findIndex((h) => h.symbol === symbol);
    if (idx < 0) return;
    const h = { ...state.holdings[idx] };
    if (profileRes.status === "fulfilled") {
      if (profileRes.value.name) h.companyName = profileRes.value.name;
      if (profileRes.value.finnhubIndustry) h.sector = profileRes.value.finnhubIndustry;
    }
    if (quoteRes.status === "fulfilled") {
      h.currentPrice = quoteRes.value.c;
      h.lastRefreshedAt = Date.now();
    }
    state.holdings[idx] = h;
    await persistAndRender();
  } catch {
    // best-effort enrichment; silent failure is OK
  }
}

/* ----------------------------- CSV import ----------------------------- */

function readFileAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read the file."));
    reader.readAsText(file);
  });
}

async function handleCsvFile(file: File): Promise<void> {
  try {
    const text = await readFileAsText(file);
    const parsed = parseCsv(text);
    if (!parsed.headers.length || !parsed.rows.length) {
      setStatus("CSV is empty or could not be parsed.", "error");
      return;
    }
    state.pendingCsv = {
      fileName: file.name,
      headers: parsed.headers,
      rows: parsed.rows,
      mapping: suggestColumnMapping(parsed.headers),
    };
    renderCsvMapper();
    setStatus(`Parsed ${parsed.rows.length} rows from ${file.name}. Review the column mapping.`, "info");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  }
}

async function confirmCsvImport(): Promise<void> {
  if (!state.pendingCsv) return;
  const { headers, rows, mapping } = state.pendingCsv;
  const result = applyMapping({ headers, rows }, mapping);
  if (result.errors.length && !result.holdings.length) {
    setStatus(result.errors[0], "error");
    return;
  }
  state.holdings = mergeHoldings(state.holdings, result.holdings);
  state.pendingCsv = null;
  renderCsvMapper();
  await persistAndRender();

  // Background enrichment for any rows missing a company name or price.
  const toEnrich = state.holdings.filter(
    (h) => !h.companyName || h.currentPrice === undefined
  );
  if (toEnrich.length) void refreshHoldings(toEnrich.map((h) => h.symbol));

  const skipMsg = result.skipped > 0 ? ` (${result.skipped} row${result.skipped === 1 ? "" : "s"} skipped)` : "";
  setStatus(`Imported ${result.holdings.length} holdings${skipMsg}.`, "success");
}

/* -------------------------- Refresh prices -------------------------- */

async function refreshAllPrices(): Promise<void> {
  if (!state.holdings.length) {
    setStatus("No holdings to refresh.", "info");
    return;
  }
  await refreshHoldings(state.holdings.map((h) => h.symbol));
}

async function refreshHoldings(symbols: string[]): Promise<void> {
  const apiKey = loadApiKey();
  if (!apiKey) {
    setStatus("Set your Finnhub API key in Settings first.", "error");
    return;
  }
  const btn = byId<HTMLButtonElement>("pf-refresh-btn");
  btn.disabled = true;
  setStatus(`Refreshing ${symbols.length} symbol${symbols.length === 1 ? "" : "s"}…`, "loading");

  let ok = 0;
  let failed = 0;
  // Small concurrency pool to avoid slamming the 60/min rate limit on free tier.
  const concurrency = 4;
  let cursor = 0;
  async function worker(): Promise<void> {
    while (cursor < symbols.length) {
      const i = cursor++;
      const symbol = symbols[i];
      try {
        const [profileRes, quoteRes] = await Promise.allSettled([
          getCompanyProfile(symbol, apiKey),
          getQuote(symbol, apiKey),
        ]);
        const idx = state.holdings.findIndex((h) => h.symbol === symbol);
        if (idx < 0) continue;
        const h = { ...state.holdings[idx] };
        if (profileRes.status === "fulfilled") {
          if (profileRes.value.name) h.companyName = profileRes.value.name;
          if (profileRes.value.finnhubIndustry) h.sector = profileRes.value.finnhubIndustry;
        }
        if (quoteRes.status === "fulfilled") {
          h.currentPrice = quoteRes.value.c;
          h.lastRefreshedAt = Date.now();
          ok++;
        } else {
          failed++;
        }
        state.holdings[idx] = h;
      } catch {
        failed++;
      }
    }
  }
  const workers: Promise<void>[] = [];
  for (let i = 0; i < Math.min(concurrency, symbols.length); i++) workers.push(worker());
  await Promise.all(workers);

  try {
    await persistAndRender();
  } finally {
    btn.disabled = false;
  }
  if (failed > 0) {
    setStatus(`Refreshed ${ok}/${symbols.length} prices. ${failed} failed (rate limit or invalid ticker).`, failed === symbols.length ? "error" : "info");
  } else {
    setStatus(`Refreshed ${ok} price${ok === 1 ? "" : "s"}.`, "success");
  }
}

/* -------------------------- Reconciliation -------------------------- */

async function startReconciliation(): Promise<void> {
  if (!state.holdings.length) {
    setStatus("Add some holdings before reconciling.", "error");
    return;
  }
  const btn = byId<HTMLButtonElement>("pf-recon-btn");
  btn.disabled = true;
  setStatus("Capturing broker selection…", "loading");
  try {
    const range = await captureBrokerSelection();
    const auto = suggestBrokerMapping(range.headers);
    state.pendingBroker = {
      range,
      mapping: auto,
    };
    renderReconPanel();
    setStatus(`Captured broker range ${range.address}. Confirm column mapping and click Run.`, "info");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
  }
}

function suggestBrokerMapping(headers: string[]): {
  symbolCol: number;
  quantityCol: number;
  costCol: number;
} {
  const m = suggestColumnMapping(headers);
  return {
    symbolCol: m.symbol,
    quantityCol: m.quantity,
    costCol: m.avgCostBasis,
  };
}

async function runReconciliation(): Promise<void> {
  if (!state.pendingBroker) return;
  const { range, mapping } = state.pendingBroker;
  if (mapping.symbolCol < 0 || mapping.quantityCol < 0) {
    setStatus("Map at least a Symbol and Quantity column.", "error");
    return;
  }
  const btn = byId<HTMLButtonElement>("pf-recon-run");
  btn.disabled = true;
  setStatus("Reconciling…", "loading");
  try {
    const rows = reconcilePortfolio(state.holdings, range, {
      qtyTolerance: 0.0001,
      costTolerance: 0.01,
      symbolCol: mapping.symbolCol,
      quantityCol: mapping.quantityCol,
      costCol: mapping.costCol,
    });
    const summary = summarizePortfolioRecon(rows);
    const sheetName = await writePortfolioReconSheet(rows, summary, range);
    state.pendingBroker = null;
    renderReconPanel();
    renderReconSummary(summary, sheetName);
    setStatus(
      `Reconciliation complete. ${summary.matched} matched · ${summary.warnings} warnings · ${summary.errors} errors. Wrote "${sheetName}".`,
      summary.errors > 0 || summary.warnings > 0 ? "info" : "success"
    );
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
  }
}

/* -------------------------- Clear --------------------------- */

/**
 * Two-click confirm: first click arms the button with a warning label for 4
 * seconds; a second click within that window actually clears. Matches the
 * lightweight `setStatus` pattern used elsewhere (no native confirm dialog).
 */
async function clearAllHoldings(): Promise<void> {
  if (!state.holdings.length) return;
  const btn = byId<HTMLButtonElement>("pf-clear-btn");
  const now = Date.now();
  if (state.clearArmedUntil < now) {
    state.clearArmedUntil = now + 4000;
    const originalText = btn.textContent;
    btn.textContent = "Click again to confirm";
    btn.classList.add("btn--danger");
    setStatus("Click Clear again within 4 seconds to wipe the portfolio.", "info");
    window.setTimeout(() => {
      if (state.clearArmedUntil && state.clearArmedUntil <= Date.now()) {
        state.clearArmedUntil = 0;
        btn.textContent = originalText;
        btn.classList.remove("btn--danger");
      }
    }, 4100);
    return;
  }
  state.clearArmedUntil = 0;
  btn.textContent = "Clear";
  btn.classList.remove("btn--danger");
  state.holdings = [];
  await persistAndRender();
  setStatus("Portfolio cleared.", "info");
}

/* --------------------------- Persistence --------------------------- */

async function persistAndRender(): Promise<void> {
  try {
    await writePortfolioToSheet(state.holdings);
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  }
  renderAll();
}

/* ----------------------------- Rendering ----------------------------- */

function renderAll(): void {
  renderTable();
  renderTotals();
  renderCsvMapper();
  renderReconPanel();
}

function renderTable(): void {
  const tbody = byId("pf-table-body");
  tbody.innerHTML = "";
  const empty = byId("pf-empty");
  if (!state.holdings.length) {
    empty.hidden = false;
    byId("pf-table-wrap").hidden = true;
    return;
  }
  empty.hidden = true;
  byId("pf-table-wrap").hidden = false;

  const totals = computeTotals(state.holdings);
  for (let i = 0; i < state.holdings.length; i++) {
    const h = state.holdings[i];
    const m = computeHoldingMetrics(h, totals.marketValue);
    const tr = document.createElement("tr");
    tr.appendChild(cell(h.symbol, "pf-td pf-td--sym"));
    const nameCell = cell(h.companyName || "—", "pf-td pf-td--name");
    // Surface sector as a hover tooltip on the company name cell.
    if (h.sector) nameCell.title = h.sector;
    tr.appendChild(nameCell);
    tr.appendChild(cell(fmtNumber(h.quantity, 4), "pf-td pf-td--num"));
    tr.appendChild(cell(fmtMoney(h.avgCostBasis, 4), "pf-td pf-td--num"));
    tr.appendChild(
      cell(
        h.currentPrice === undefined ? "—" : fmtMoney(h.currentPrice, 2),
        "pf-td pf-td--num"
      )
    );
    tr.appendChild(cell(fmtMoney(m.marketValue, 2), "pf-td pf-td--num"));
    tr.appendChild(
      cell(
        fmtSignedMoney(m.unrealizedPnl),
        "pf-td pf-td--num " + pnlClass(m.unrealizedPnl)
      )
    );
    tr.appendChild(
      cell(
        fmtSignedPct(m.unrealizedPnlPct),
        "pf-td pf-td--num " + pnlClass(m.unrealizedPnl)
      )
    );
    tr.appendChild(cell(fmtPct(m.weightPct), "pf-td pf-td--num"));
    const actions = document.createElement("td");
    actions.className = "pf-td pf-td--actions";
    const del = document.createElement("button");
    del.type = "button";
    del.className = "pf-row-del";
    del.title = `Remove ${h.symbol}`;
    del.setAttribute("aria-label", `Remove ${h.symbol}`);
    del.textContent = "×";
    del.addEventListener("click", () => void removeHolding(h.symbol));
    actions.appendChild(del);
    tr.appendChild(actions);
    tbody.appendChild(tr);
  }
}

async function removeHolding(symbol: string): Promise<void> {
  state.holdings = state.holdings.filter((h) => h.symbol !== symbol);
  await persistAndRender();
  setStatus(`Removed ${symbol}.`, "info");
}

function renderTotals(): void {
  const totals = computeTotals(state.holdings);
  byId("pf-total-market").textContent = fmtMoney(totals.marketValue, 2);
  byId("pf-total-cost").textContent = fmtMoney(totals.costTotal, 2);
  const pnlEl = byId("pf-total-pnl");
  pnlEl.textContent = fmtSignedMoney(totals.unrealizedPnl);
  pnlEl.className = "pf-metric__value " + pnlClass(totals.unrealizedPnl);
  const pctEl = byId("pf-total-pnl-pct");
  pctEl.textContent = fmtSignedPct(totals.unrealizedPnlPct);
  pctEl.className = "pf-metric__value pf-metric__value--small " + pnlClass(totals.unrealizedPnl);
  byId("pf-total-rows").textContent = `${totals.rowCount} holding${totals.rowCount === 1 ? "" : "s"}`;
}

/* ------------------------- CSV mapper render ------------------------- */

function renderCsvMapper(): void {
  const panel = byId("pf-csv-panel");
  if (!state.pendingCsv) {
    panel.hidden = true;
    return;
  }
  panel.hidden = false;
  byId("pf-csv-filename").textContent = state.pendingCsv.fileName;
  byId("pf-csv-count").textContent = `${state.pendingCsv.rows.length} rows`;

  const mapHost = byId("pf-csv-mapping");
  mapHost.innerHTML = "";
  for (let i = 0; i < HOLDING_FIELDS.length; i++) {
    const field = HOLDING_FIELDS[i];
    const required = REQUIRED_IMPORT_FIELDS.indexOf(field) >= 0;
    mapHost.appendChild(buildMappingRow(field, required));
  }
  renderCsvPreview();
}

function buildMappingRow(field: HoldingField, required: boolean): HTMLElement {
  if (!state.pendingCsv) return document.createElement("div");
  const wrap = document.createElement("label");
  wrap.className = "pf-map-row";
  const label = document.createElement("span");
  label.className = "pf-map-label";
  label.textContent = HOLDING_FIELD_LABELS[field] + (required ? " *" : "");
  wrap.appendChild(label);

  const select = document.createElement("select");
  select.className = "field__input";
  const unset = document.createElement("option");
  unset.value = "-1";
  unset.textContent = "— unmapped —";
  select.appendChild(unset);
  for (let c = 0; c < state.pendingCsv.headers.length; c++) {
    const opt = document.createElement("option");
    opt.value = String(c);
    const h = state.pendingCsv.headers[c];
    opt.textContent = h ? `${columnLetter(c)} — ${h}` : `Column ${columnLetter(c)}`;
    select.appendChild(opt);
  }
  select.value = String(state.pendingCsv.mapping[field] ?? -1);
  select.addEventListener("change", () => {
    if (!state.pendingCsv) return;
    state.pendingCsv.mapping[field] = parseInt(select.value, 10);
    renderCsvPreview();
  });
  wrap.appendChild(select);
  return wrap;
}

function renderCsvPreview(): void {
  const host = byId("pf-csv-preview");
  host.innerHTML = "";
  if (!state.pendingCsv) return;
  const { rows, mapping } = state.pendingCsv;
  const sample = rows.slice(0, 3);
  if (!sample.length) return;
  const previewTable = document.createElement("table");
  previewTable.className = "pf-preview-table";
  const thead = document.createElement("thead");
  const htr = document.createElement("tr");
  for (let i = 0; i < HOLDING_FIELDS.length; i++) {
    const th = document.createElement("th");
    th.textContent = HOLDING_FIELD_LABELS[HOLDING_FIELDS[i]];
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  previewTable.appendChild(thead);
  const tbody = document.createElement("tbody");
  for (let r = 0; r < sample.length; r++) {
    const tr = document.createElement("tr");
    for (let f = 0; f < HOLDING_FIELDS.length; f++) {
      const td = document.createElement("td");
      const col = mapping[HOLDING_FIELDS[f]];
      td.textContent = col >= 0 && col < sample[r].length ? sample[r][col] || "" : "";
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  previewTable.appendChild(tbody);
  host.appendChild(previewTable);
}

/* ----------------------- Reconciliation render ----------------------- */

function renderReconPanel(): void {
  const panel = byId("pf-recon-panel");
  if (!state.pendingBroker) {
    panel.hidden = true;
    return;
  }
  panel.hidden = false;
  byId("pf-recon-range").textContent = `${state.pendingBroker.range.address} · ${state.pendingBroker.range.rows.length} data rows`;
  const host = byId("pf-recon-mapping");
  host.innerHTML = "";
  const rows: Array<{ key: "symbolCol" | "quantityCol" | "costCol"; label: string; required: boolean }> = [
    { key: "symbolCol", label: "Symbol", required: true },
    { key: "quantityCol", label: "Quantity", required: true },
    { key: "costCol", label: "Cost Basis", required: false },
  ];
  const headers = state.pendingBroker.range.headers;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const wrap = document.createElement("label");
    wrap.className = "pf-map-row";
    const label = document.createElement("span");
    label.className = "pf-map-label";
    label.textContent = row.label + (row.required ? " *" : "");
    wrap.appendChild(label);
    const select = document.createElement("select");
    select.className = "field__input";
    const unset = document.createElement("option");
    unset.value = "-1";
    unset.textContent = "— unmapped —";
    select.appendChild(unset);
    for (let c = 0; c < headers.length; c++) {
      const opt = document.createElement("option");
      opt.value = String(c);
      opt.textContent = headers[c] ? `${columnLetter(c)} — ${headers[c]}` : `Column ${columnLetter(c)}`;
      select.appendChild(opt);
    }
    select.value = String(state.pendingBroker!.mapping[row.key]);
    select.addEventListener("change", () => {
      if (!state.pendingBroker) return;
      state.pendingBroker.mapping[row.key] = parseInt(select.value, 10);
    });
    wrap.appendChild(select);
    host.appendChild(wrap);
  }
}

function renderReconSummary(
  summary: { matched: number; warnings: number; errors: number; total: number },
  sheetName: string
): void {
  const el = byId("pf-recon-summary");
  el.hidden = false;
  el.innerHTML = `
    <h4>Last reconciliation</h4>
    <div class="pf-recon-grid">
      <div class="pf-recon-stat pf-recon-stat--ok"><div class="pf-recon-stat__v">${summary.matched}</div><div class="pf-recon-stat__l">Matched</div></div>
      <div class="pf-recon-stat pf-recon-stat--warn"><div class="pf-recon-stat__v">${summary.warnings}</div><div class="pf-recon-stat__l">Warnings</div></div>
      <div class="pf-recon-stat pf-recon-stat--err"><div class="pf-recon-stat__v">${summary.errors}</div><div class="pf-recon-stat__l">Errors</div></div>
      <div class="pf-recon-stat"><div class="pf-recon-stat__v">${summary.total}</div><div class="pf-recon-stat__l">Total</div></div>
    </div>
    <p class="pf-recon-foot">Results written to "${escapeHtml(sheetName)}".</p>
  `;
}

/* --------------------------- Format helpers --------------------------- */

function cell(text: string, className: string): HTMLElement {
  const td = document.createElement("td");
  td.className = className;
  td.textContent = text;
  return td;
}

function fmtNumber(n: number, decimals: number): string {
  if (!isFinite(n)) return "—";
  return n.toLocaleString(undefined, {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
}

function fmtMoney(n: number, decimals: number): string {
  if (!isFinite(n)) return "—";
  return (
    "$" +
    n.toLocaleString(undefined, {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    })
  );
}

function fmtSignedMoney(n: number): string {
  if (!isFinite(n)) return "—";
  const sign = n > 0 ? "+" : n < 0 ? "−" : "";
  const abs = Math.abs(n);
  return (
    sign +
    "$" +
    abs.toLocaleString(undefined, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })
  );
}

function fmtPct(n: number): string {
  if (!isFinite(n)) return "—";
  return n.toFixed(2) + "%";
}

function fmtSignedPct(n: number): string {
  if (!isFinite(n)) return "—";
  const sign = n > 0 ? "+" : n < 0 ? "−" : "";
  return sign + Math.abs(n).toFixed(2) + "%";
}

function pnlClass(n: number): string {
  if (n > 0) return "pf-pnl--up";
  if (n < 0) return "pf-pnl--down";
  return "";
}

function columnLetter(i: number): string {
  let s = "";
  let n = i;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => {
    switch (c) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      case "'":
        return "&#39;";
      default:
        return c;
    }
  });
}
