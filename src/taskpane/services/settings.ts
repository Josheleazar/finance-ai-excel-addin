/*
 * API key storage.
 *
 * Primary: Office.context.document.settings (persisted per workbook, survives reopen).
 * Fallback: window.localStorage (per browser/origin; used if document settings unavailable).
 *
 * Neither of these is "secure" in the cryptographic sense. For production, prefer
 * OfficeRuntime.storage (shared runtime) or an auth flow that keeps the key server-side.
 */

/* global Office, window */

import { MatchingMode } from "./reconciliation";

const SETTING_KEY = "finalysis.finnhubApiKey";
const LOCAL_STORAGE_KEY = "finalysis.finnhubApiKey";
const RECON_SETTING_KEY = "finalysis.reconDefaults";
const RECON_LOCAL_STORAGE_KEY = "finalysis.reconDefaults";

function hasDocumentSettings(): boolean {
  try {
    return !!(
      Office &&
      Office.context &&
      Office.context.document &&
      Office.context.document.settings
    );
  } catch {
    return false;
  }
}

function saveDocumentSettingsAsync(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error ? result.error.message : "Failed to save settings"));
      }
    });
  });
}

export function loadApiKey(): string {
  if (hasDocumentSettings()) {
    const value = Office.context.document.settings.get(SETTING_KEY);
    if (typeof value === "string" && value) {
      return value;
    }
  }
  try {
    return window.localStorage.getItem(LOCAL_STORAGE_KEY) || "";
  } catch {
    return "";
  }
}

export async function saveApiKey(apiKey: string): Promise<void> {
  const trimmed = apiKey.trim();

  // Always mirror to localStorage as a fallback.
  try {
    if (trimmed) {
      window.localStorage.setItem(LOCAL_STORAGE_KEY, trimmed);
    } else {
      window.localStorage.removeItem(LOCAL_STORAGE_KEY);
    }
  } catch {
    // localStorage can be disabled; ignore.
  }

  if (hasDocumentSettings()) {
    if (trimmed) {
      Office.context.document.settings.set(SETTING_KEY, trimmed);
    } else {
      Office.context.document.settings.remove(SETTING_KEY);
    }
    await saveDocumentSettingsAsync();
  }
}

export function maskApiKey(apiKey: string): string {
  if (!apiKey) return "";
  if (apiKey.length <= 6) return "*".repeat(apiKey.length);
  return `${apiKey.slice(0, 3)}${"*".repeat(Math.max(4, apiKey.length - 6))}${apiKey.slice(-3)}`;
}

/* -------------------- Reconciliation defaults -------------------- */

export interface ReconSettings {
  mode: MatchingMode;
  amountFixed: number;
  amountPct: number;
  dateDays: number;
}

export const DEFAULT_RECON_SETTINGS: ReconSettings = {
  mode: "normal",
  amountFixed: 0.01,
  amountPct: 0,
  dateDays: 1,
};

function coerceReconSettings(raw: unknown): ReconSettings {
  const src = (raw && typeof raw === "object" ? raw : {}) as { [k: string]: unknown };
  const mode = src.mode;
  const allowedModes: MatchingMode[] = ["strict", "normal", "loose", "custom"];
  return {
    mode:
      typeof mode === "string" && allowedModes.indexOf(mode as MatchingMode) >= 0
        ? (mode as MatchingMode)
        : DEFAULT_RECON_SETTINGS.mode,
    amountFixed:
      typeof src.amountFixed === "number" && isFinite(src.amountFixed) && src.amountFixed >= 0
        ? src.amountFixed
        : DEFAULT_RECON_SETTINGS.amountFixed,
    amountPct:
      typeof src.amountPct === "number" && isFinite(src.amountPct) && src.amountPct >= 0
        ? src.amountPct
        : DEFAULT_RECON_SETTINGS.amountPct,
    dateDays:
      typeof src.dateDays === "number" && isFinite(src.dateDays) && src.dateDays >= 0
        ? Math.floor(src.dateDays)
        : DEFAULT_RECON_SETTINGS.dateDays,
  };
}

export function loadReconSettings(): ReconSettings {
  if (hasDocumentSettings()) {
    const value = Office.context.document.settings.get(RECON_SETTING_KEY);
    if (value && typeof value === "object") {
      return coerceReconSettings(value);
    }
  }
  try {
    const raw = window.localStorage.getItem(RECON_LOCAL_STORAGE_KEY);
    if (raw) {
      return coerceReconSettings(JSON.parse(raw));
    }
  } catch {
    // localStorage disabled or JSON invalid — fall back to defaults.
  }
  return { ...DEFAULT_RECON_SETTINGS };
}

export async function saveReconSettings(settings: ReconSettings): Promise<void> {
  const normalized = coerceReconSettings(settings);
  try {
    window.localStorage.setItem(RECON_LOCAL_STORAGE_KEY, JSON.stringify(normalized));
  } catch {
    // localStorage can be disabled; ignore.
  }
  if (hasDocumentSettings()) {
    Office.context.document.settings.set(RECON_SETTING_KEY, normalized);
    await saveDocumentSettingsAsync();
  }
}
