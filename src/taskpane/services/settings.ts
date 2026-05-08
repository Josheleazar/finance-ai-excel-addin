/*
 * Settings persistence.
 *
 *   - API key (Finnhub).
 *   - Reconciliation defaults: sensitivity + dynamic field configuration.
 *
 * Primary store: Office.context.document.settings (workbook-scoped, survives reopen).
 * Fallback: window.localStorage (per browser/origin). Neither is cryptographically
 * secure; for production prefer a server-side proxy or OfficeRuntime.storage.
 */

/* global Office, window */

import {
  FieldTolerance,
  FieldType,
  FieldWeight,
  MAX_FIELDS,
  Sensitivity,
  defaultToleranceForType,
} from "./reconciliation";

const SETTING_KEY = "finalysis.finnhubApiKey";
const LOCAL_STORAGE_KEY = "finalysis.finnhubApiKey";
// Bumped to v2 when the reconciliation model switched from fixed
// Date/Amount/Description to dynamic fields — incompatible shape.
const RECON_SETTING_KEY = "finalysis.reconDefaults.v2";
const RECON_LOCAL_STORAGE_KEY = "finalysis.reconDefaults.v2";

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

/**
 * The persisted shape of a single field. Omits transient column indexes
 * (colA/colB) since those are tied to a specific captured range.
 */
export interface PersistedField {
  id: string;
  label: string;
  type: FieldType;
  weight: FieldWeight;
  required: boolean;
  downgradeOnFail: boolean;
  tolerance: FieldTolerance;
}

export interface ReconSettings {
  sensitivity: Sensitivity;
  fields: PersistedField[];
}

/** Sensible starting configuration for a transactional bank/ledger reconciliation. */
export const DEFAULT_RECON_SETTINGS: ReconSettings = {
  sensitivity: "normal",
  fields: [
    {
      id: "f1",
      label: "Date",
      type: "date",
      weight: "medium",
      required: false,
      downgradeOnFail: false,
      tolerance: defaultToleranceForType("date"),
    },
    {
      id: "f2",
      label: "Amount",
      type: "numeric",
      weight: "high",
      required: true,
      downgradeOnFail: false,
      tolerance: defaultToleranceForType("numeric"),
    },
    {
      id: "f3",
      label: "Description",
      type: "fuzzy",
      weight: "medium",
      required: false,
      downgradeOnFail: false,
      tolerance: defaultToleranceForType("fuzzy"),
    },
  ],
};

const ALLOWED_TYPES: FieldType[] = ["exact", "numeric", "date", "fuzzy"];
const ALLOWED_WEIGHTS: FieldWeight[] = ["low", "medium", "high"];
const ALLOWED_SENSITIVITIES: Sensitivity[] = ["strict", "normal", "loose"];

function coerceNonNegative(v: unknown, fallback: number): number {
  return typeof v === "number" && isFinite(v) && v >= 0 ? v : fallback;
}

function coerceTolerance(raw: unknown, type: FieldType): FieldTolerance {
  const def = defaultToleranceForType(type);
  const src = raw && typeof raw === "object" ? (raw as { [k: string]: unknown }) : {};
  const t: FieldTolerance = {};
  if (type === "numeric") {
    t.amountFixed = coerceNonNegative(src.amountFixed, def.amountFixed || 0);
    t.amountPct = coerceNonNegative(src.amountPct, def.amountPct || 0);
  } else if (type === "date") {
    t.dateDays = Math.floor(coerceNonNegative(src.dateDays, def.dateDays || 0));
  } else if (type === "fuzzy") {
    const n = coerceNonNegative(src.minSimilarity, def.minSimilarity || 0);
    t.minSimilarity = Math.max(0, Math.min(1, n));
  } else if (type === "exact") {
    t.caseSensitive = !!src.caseSensitive;
  }
  return t;
}

function coerceField(raw: unknown, i: number): PersistedField | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as { [k: string]: unknown };
  const type =
    typeof r.type === "string" && ALLOWED_TYPES.indexOf(r.type as FieldType) >= 0
      ? (r.type as FieldType)
      : null;
  if (!type) return null;
  const weight =
    typeof r.weight === "string" && ALLOWED_WEIGHTS.indexOf(r.weight as FieldWeight) >= 0
      ? (r.weight as FieldWeight)
      : "medium";
  return {
    id: typeof r.id === "string" && r.id ? r.id : `f${i + 1}`,
    label: typeof r.label === "string" && r.label ? r.label : `Field ${i + 1}`,
    type,
    weight,
    required: !!r.required,
    downgradeOnFail: !!r.downgradeOnFail,
    tolerance: coerceTolerance(r.tolerance, type),
  };
}

function coerceReconSettings(raw: unknown): ReconSettings {
  const src = (raw && typeof raw === "object" ? raw : {}) as { [k: string]: unknown };
  const sensitivity =
    typeof src.sensitivity === "string" &&
    ALLOWED_SENSITIVITIES.indexOf(src.sensitivity as Sensitivity) >= 0
      ? (src.sensitivity as Sensitivity)
      : DEFAULT_RECON_SETTINGS.sensitivity;

  let fields: PersistedField[] = [];
  if (Array.isArray(src.fields)) {
    for (let i = 0; i < src.fields.length && fields.length < MAX_FIELDS; i++) {
      const f = coerceField(src.fields[i], fields.length);
      if (f) fields.push(f);
    }
  }
  if (!fields.length) {
    // No saved / invalid fields — fall back to the baked-in defaults.
    fields = DEFAULT_RECON_SETTINGS.fields.map((f) => ({
      ...f,
      tolerance: { ...f.tolerance },
    }));
  }
  return { sensitivity, fields };
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
    // localStorage disabled or JSON invalid — fall through to defaults.
  }
  return {
    sensitivity: DEFAULT_RECON_SETTINGS.sensitivity,
    fields: DEFAULT_RECON_SETTINGS.fields.map((f) => ({
      ...f,
      tolerance: { ...f.tolerance },
    })),
  };
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
