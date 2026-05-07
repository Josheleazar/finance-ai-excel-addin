/*
 * Shared UI helpers. Kept tiny on purpose; tabs import from here rather than
 * reaching into taskpane.ts (which would create a circular import).
 */

/* global document, HTMLElement, HTMLInputElement */

export type StatusKind = "info" | "success" | "error" | "loading";

export function setStatus(message: string, kind: StatusKind = "info"): void {
  const el = document.getElementById("status-strip");
  if (!el) return;
  el.textContent = message;
  el.dataset.kind = kind;
  el.classList.toggle("is-empty", !message);
}

export function clearStatus(): void {
  setStatus("", "info");
}

export function byId<T extends HTMLElement = HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`Element #${id} not found`);
  }
  return el as T;
}

export function inputValue(id: string): string {
  return byId<HTMLInputElement>(id).value.trim();
}

export function formatUnix(ts: number): string {
  if (!ts) return "";
  const d = new Date(ts * 1000);
  return d.toISOString().replace("T", " ").slice(0, 19);
}
