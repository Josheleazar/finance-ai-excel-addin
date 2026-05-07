/*
 * Finnhub API client.
 *
 * NOTE: Calls go directly from the task pane to finnhub.io using the user's API key.
 * For production, route through a backend proxy so the key never reaches the client.
 * Swap FINNHUB_BASE_URL to your proxy origin when ready.
 *
 * Tier caveat: /stock/candle is a premium endpoint on Finnhub's current pricing.
 * Free-tier keys will receive 403 Forbidden. Upgrade the key or swap in an
 * equivalent provider to use the candles feature.
 */

/* global URLSearchParams, fetch */

import {
  BasicFinancials,
  CandleResolution,
  CandleResponse,
  CompanyProfile,
  Quote,
} from "../types/finnhub";

// TODO(security): replace with your backend proxy URL, e.g. "https://api.example.com/finnhub".
export const FINNHUB_BASE_URL = "https://finnhub.io/api/v1";

function buildUrl(
  path: string,
  params: { [key: string]: string | number },
  apiKey: string
): string {
  const query = new URLSearchParams();
  Object.keys(params).forEach((k) => {
    const v = params[k];
    if (v !== undefined && v !== null) {
      query.append(k, String(v));
    }
  });
  query.append("token", apiKey);
  return `${FINNHUB_BASE_URL}${path}?${query.toString()}`;
}

async function request<T>(url: string, context?: string): Promise<T> {
  const res = await fetch(url);
  if (!res.ok) {
    let detail = res.statusText;
    try {
      const body = await res.text();
      if (body) detail = body;
    } catch {
      // ignore body parse errors
    }
    // Finnhub returns 403 for endpoints not available on the user's plan
    // (e.g. /stock/candle is premium-only). Give a friendlier hint.
    if (res.status === 403 && context === "candles") {
      throw new Error(
        "Historical candles require a paid Finnhub plan (the free tier returns 403). " +
          "Upgrade your key or use a different data source."
      );
    }
    if (res.status === 401) {
      throw new Error("Invalid Finnhub API key. Open settings to update it.");
    }
    if (res.status === 429) {
      throw new Error(
        "Finnhub rate limit hit (60 calls/min on free tier). Wait a moment and retry."
      );
    }
    throw new Error(`Finnhub ${res.status}: ${detail}`);
  }
  return (await res.json()) as T;
}

function requireKey(apiKey: string): void {
  if (!apiKey || !apiKey.trim()) {
    throw new Error("Finnhub API key is not set. Open settings to add one.");
  }
}

export async function getQuote(symbol: string, apiKey: string): Promise<Quote> {
  requireKey(apiKey);
  const url = buildUrl("/quote", { symbol: symbol.toUpperCase() }, apiKey);
  const data = await request<Quote>(url);
  if (!data || typeof data.c !== "number" || data.c === 0) {
    throw new Error(`No quote data for "${symbol}". Check the ticker symbol.`);
  }
  return data;
}

export async function getCandles(
  symbol: string,
  resolution: CandleResolution,
  fromUnix: number,
  toUnix: number,
  apiKey: string
): Promise<CandleResponse> {
  requireKey(apiKey);
  const url = buildUrl(
    "/stock/candle",
    { symbol: symbol.toUpperCase(), resolution, from: fromUnix, to: toUnix },
    apiKey
  );
  const data = await request<CandleResponse>(url, "candles");
  if (data.s !== "ok") {
    throw new Error(`No candle data for "${symbol}" in the requested range.`);
  }
  return data;
}

export async function getCompanyProfile(symbol: string, apiKey: string): Promise<CompanyProfile> {
  requireKey(apiKey);
  const url = buildUrl("/stock/profile2", { symbol: symbol.toUpperCase() }, apiKey);
  const data = await request<CompanyProfile>(url);
  if (!data || !data.name) {
    throw new Error(`No profile found for "${symbol}".`);
  }
  return data;
}

export async function getBasicFinancials(
  symbol: string,
  apiKey: string,
  metric: string = "all"
): Promise<BasicFinancials> {
  requireKey(apiKey);
  const url = buildUrl("/stock/metric", { symbol: symbol.toUpperCase(), metric }, apiKey);
  const data = await request<BasicFinancials>(url);
  if (!data || !data.metric) {
    throw new Error(`No financial metrics for "${symbol}".`);
  }
  return data;
}

/** Convert a lookback in days to a [fromUnix, toUnix] pair ending at now. */
export function rangeFromDaysAgo(days: number): { fromUnix: number; toUnix: number } {
  const now = Math.floor(Date.now() / 1000);
  return { fromUnix: now - days * 24 * 60 * 60, toUnix: now };
}
