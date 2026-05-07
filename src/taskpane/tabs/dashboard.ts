/*
 * Dashboard tab: quick in-pane quote lookup (no sheet write).
 * Useful for verifying the API key and checking a ticker before importing.
 */

/* global HTMLInputElement, HTMLButtonElement, KeyboardEvent */

import { getQuote } from "../services/finnhub";
import { loadApiKey } from "../services/settings";
import { byId, formatUnix, inputValue, setStatus } from "../services/ui";
import { Quote } from "../types/finnhub";

export function initDashboardTab(): void {
  byId("dash-lookup-btn").addEventListener("click", () => {
    void handleLookup();
  });

  byId<HTMLInputElement>("dash-symbol").addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key === "Enter") {
      void handleLookup();
    }
  });
}

async function handleLookup(): Promise<void> {
  const symbol = inputValue("dash-symbol");
  if (!symbol) {
    setStatus("Enter a ticker symbol first.", "error");
    return;
  }

  const apiKey = loadApiKey();
  const btn = byId<HTMLButtonElement>("dash-lookup-btn");
  btn.disabled = true;
  setStatus(`Fetching ${symbol.toUpperCase()}\u2026`, "loading");

  try {
    const quote = await getQuote(symbol, apiKey);
    renderQuote(symbol.toUpperCase(), quote);
    setStatus(`Loaded ${symbol.toUpperCase()}.`, "success");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
  }
}

function renderQuote(symbol: string, q: Quote): void {
  const container = byId("dash-quote-card");
  const changeAbs = typeof q.d === "number" ? q.d : q.c - q.pc;
  const changePct = typeof q.dp === "number" ? q.dp : q.pc ? ((q.c - q.pc) / q.pc) * 100 : 0;
  const direction = changeAbs >= 0 ? "up" : "down";
  const sign = changeAbs >= 0 ? "+" : "";

  container.innerHTML = `
    <div class="quote-card__header">
      <span class="quote-card__symbol">${escapeHtml(symbol)}</span>
      <span class="quote-card__time">${escapeHtml(formatUnix(q.t))}</span>
    </div>
    <div class="quote-card__price">${q.c.toFixed(2)}</div>
    <div class="quote-card__change quote-card__change--${direction}">
      ${sign}${changeAbs.toFixed(2)} (${sign}${changePct.toFixed(2)}%)
    </div>
    <dl class="quote-card__grid">
      <div><dt>Open</dt><dd>${q.o.toFixed(2)}</dd></div>
      <div><dt>High</dt><dd>${q.h.toFixed(2)}</dd></div>
      <div><dt>Low</dt><dd>${q.l.toFixed(2)}</dd></div>
      <div><dt>Prev Close</dt><dd>${q.pc.toFixed(2)}</dd></div>
    </dl>
  `;
  container.classList.remove("is-empty");
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
