/*
 * Typed shapes for Finnhub REST responses.
 * Reference: https://finnhub.io/docs/api
 */

export interface Quote {
  c: number; // current price
  h: number; // high of the day
  l: number; // low of the day
  o: number; // open of the day
  pc: number; // previous close
  t: number; // timestamp (unix seconds)
  d?: number; // change (absolute)
  dp?: number; // change (percent)
}

export type CandleResolution = "1" | "5" | "15" | "30" | "60" | "D" | "W" | "M";

export interface CandleResponse {
  s: "ok" | "no_data";
  c: number[];
  h: number[];
  l: number[];
  o: number[];
  t: number[];
  v: number[];
}

export interface CompanyProfile {
  country?: string;
  currency?: string;
  exchange?: string;
  finnhubIndustry?: string;
  ipo?: string;
  logo?: string;
  marketCapitalization?: number;
  name?: string;
  phone?: string;
  shareOutstanding?: number;
  ticker?: string;
  weburl?: string;
}

export interface BasicFinancials {
  symbol: string;
  metricType: string;
  metric: { [key: string]: number | string | undefined };
  series?: {
    annual?: { [key: string]: Array<{ period: string; v: number }> };
    quarterly?: { [key: string]: Array<{ period: string; v: number }> };
  };
}
