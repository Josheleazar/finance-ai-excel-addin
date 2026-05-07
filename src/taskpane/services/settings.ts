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

const SETTING_KEY = "finalysis.finnhubApiKey";
const LOCAL_STORAGE_KEY = "finalysis.finnhubApiKey";

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
