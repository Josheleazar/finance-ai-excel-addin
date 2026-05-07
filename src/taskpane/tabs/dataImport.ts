/*
 * Data Import tab: fetches from Finnhub and writes results to a new worksheet.
 *
 * Operations:
 *   - Quote (current)
 *   - Historical candles (daily/weekly/monthly, configurable lookback)
 *   - Company profile
 *   - Basic financials (metric=all)
 */

/* global HTMLButtonElement, HTMLSelectElement */

import {
  getBasicFinancials,
  getCandles,
  getCompanyProfile,
  getQuote,
  rangeFromDaysAgo,
} from "../services/finnhub";
import { loadApiKey } from "../services/settings";
import { writeKeyValueToNewSheet, writeTableToNewSheet } from "../services/excel";
import { byId, formatUnix, inputValue, setStatus } from "../services/ui";
import { CandleResolution } from "../types/finnhub";

export function initDataImportTab(): void {
  byId("di-quote-btn").addEventListener("click", () => void runQuote());
  byId("di-candles-btn").addEventListener("click", () => void runCandles());
  byId("di-profile-btn").addEventListener("click", () => void runProfile());
  byId("di-financials-btn").addEventListener("click", () => void runFinancials());
}

function currentSymbol(): string {
  const s = inputValue("di-symbol");
  if (!s) throw new Error("Enter a ticker symbol first.");
  return s.toUpperCase();
}

async function withBusy<T>(buttonId: string, label: string, fn: () => Promise<T>): Promise<void> {
  const btn = byId<HTMLButtonElement>(buttonId);
  btn.disabled = true;
  setStatus(`${label}\u2026`, "loading");
  try {
    await fn();
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    btn.disabled = false;
  }
}

async function runQuote(): Promise<void> {
  await withBusy("di-quote-btn", "Fetching quote", async () => {
    const symbol = currentSymbol();
    const apiKey = loadApiKey();
    const q = await getQuote(symbol, apiKey);
    const sheet = await writeKeyValueToNewSheet(`${symbol} Quote`, `${symbol} \u2014 Quote`, [
      ["Symbol", symbol],
      ["Current", q.c],
      ["Open", q.o],
      ["High", q.h],
      ["Low", q.l],
      ["Previous Close", q.pc],
      ["Change", typeof q.d === "number" ? q.d : q.c - q.pc],
      ["Change %", typeof q.dp === "number" ? q.dp : q.pc ? ((q.c - q.pc) / q.pc) * 100 : 0],
      ["As of (UTC)", formatUnix(q.t)],
    ]);
    setStatus(`Wrote quote to "${sheet}".`, "success");
  });
}

async function runCandles(): Promise<void> {
  await withBusy("di-candles-btn", "Fetching candles", async () => {
    const symbol = currentSymbol();
    const apiKey = loadApiKey();
    const resolution = byId<HTMLSelectElement>("di-candles-resolution").value as CandleResolution;
    const days = parseInt(inputValue("di-candles-days"), 10);
    if (!days || days <= 0) {
      throw new Error("Enter a positive number of days.");
    }
    const { fromUnix, toUnix } = rangeFromDaysAgo(days);
    const data = await getCandles(symbol, resolution, fromUnix, toUnix, apiKey);

    const rows = data.t.map((ts, i) => [
      formatUnix(ts),
      data.o[i],
      data.h[i],
      data.l[i],
      data.c[i],
      data.v[i],
    ]);

    const sheet = await writeTableToNewSheet(
      `${symbol} Candles`,
      ["Date (UTC)", "Open", "High", "Low", "Close", "Volume"],
      rows
    );
    setStatus(`Wrote ${rows.length} candles to "${sheet}".`, "success");
  });
}

async function runProfile(): Promise<void> {
  await withBusy("di-profile-btn", "Fetching profile", async () => {
    const symbol = currentSymbol();
    const apiKey = loadApiKey();
    const p = await getCompanyProfile(symbol, apiKey);
    const sheet = await writeKeyValueToNewSheet(
      `${symbol} Profile`,
      `${symbol} \u2014 Company Profile`,
      [
        ["Name", p.name || ""],
        ["Ticker", p.ticker || symbol],
        ["Exchange", p.exchange || ""],
        ["Industry", p.finnhubIndustry || ""],
        ["Country", p.country || ""],
        ["Currency", p.currency || ""],
        ["IPO", p.ipo || ""],
        ["Market Cap (M)", p.marketCapitalization ?? ""],
        ["Shares Outstanding (M)", p.shareOutstanding ?? ""],
        ["Phone", p.phone || ""],
        ["Website", p.weburl || ""],
        ["Logo", p.logo || ""],
      ]
    );
    setStatus(`Wrote profile to "${sheet}".`, "success");
  });
}

async function runFinancials(): Promise<void> {
  await withBusy("di-financials-btn", "Fetching financial metrics", async () => {
    const symbol = currentSymbol();
    const apiKey = loadApiKey();
    const data = await getBasicFinancials(symbol, apiKey);
    const entries = Object.keys(data.metric)
      .sort()
      .map((k) => {
        const v = data.metric[k];
        return [k, (v === undefined ? "" : v) as string | number];
      }) as Array<[string, string | number]>;

    const sheet = await writeTableToNewSheet(`${symbol} Financials`, ["Metric", "Value"], entries);
    setStatus(`Wrote ${entries.length} metrics to "${sheet}".`, "success");
  });
}
