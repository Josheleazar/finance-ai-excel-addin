/*
 * Finalysis task pane bootstrap.
 *
 * Responsibilities:
 *   - Wait for Office.onReady and verify the host is Excel.
 *   - Gate the UI behind Clerk authentication. If no JWT is stored, render
 *     the sign-in gate; only once the dialog returns a valid token do we
 *     hydrate tabs.
 *   - Hydrate the settings panel (API key) from storage.
 *   - Wire up the tab switcher.
 *   - Delegate per-tab event handlers to their modules.
 */

/* global console, document, Office, HTMLElement, HTMLButtonElement, HTMLInputElement */

import { initDashboardTab } from "./tabs/dashboard";
import { initDataImportTab } from "./tabs/dataImport";
import { initReconciliationTab } from "./tabs/reconciliation";
import { initPortfolioTab } from "./tabs/portfolio";
import {
  AuthState,
  getStoredAuth,
  onAuthStateChange,
  openSignInDialog,
  signOut,
} from "./services/auth";
import { loadApiKey, maskApiKey, saveApiKey } from "./services/settings";
import { byId, clearStatus, setStatus } from "./services/ui";

let tabsBootstrapped = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    showSideloadMessage();
    return;
  }

  void (async () => {
    try {
      await initAuthGate();
    } catch (err) {
      console.error("Finalysis bootstrap failed", err);
      setStatus(err instanceof Error ? err.message : "Failed to start the add-in.", "error");
    }
  })();
});

/* --- Auth gate --- */

async function initAuthGate(): Promise<void> {
  byId("app-root").classList.remove("is-loading");

  const signInBtn = byId<HTMLButtonElement>("auth-signin-btn");
  const signOutBtn = byId<HTMLButtonElement>("auth-signout");

  signInBtn.addEventListener("click", () => {
    void handleSignIn();
  });

  signOutBtn.addEventListener("click", () => {
    void handleSignOut();
  });

  // React to programmatic auth changes (e.g. a 401 elsewhere clearing state).
  onAuthStateChange((state) => {
    renderAuthState(state);
  });

  const initial = await getStoredAuth();
  renderAuthState(initial);

  // If we already have a stored session, run bootstrap immediately. Otherwise
  // wait for the user to click "Sign in".
  if (initial) {
    bootstrapOnce();
  }
}

async function handleSignIn(): Promise<void> {
  const errEl = byId("auth-error");
  errEl.hidden = true;
  errEl.textContent = "";
  const btn = byId<HTMLButtonElement>("auth-signin-btn");
  btn.disabled = true;
  const originalLabel = btn.textContent;
  btn.textContent = "Opening sign-in…";
  try {
    const state = await openSignInDialog();
    renderAuthState(state);
    bootstrapOnce();
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    errEl.textContent = message;
    errEl.hidden = false;
  } finally {
    btn.disabled = false;
    btn.textContent = originalLabel || "Sign in";
  }
}

async function handleSignOut(): Promise<void> {
  try {
    await signOut();
    renderAuthState(null);
    // We don't tear down the already-bootstrapped tabs (their event listeners
    // are harmless) — we just hide them behind the gate. The user can sign
    // back in to reveal them again.
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  }
}

function renderAuthState(state: AuthState | null): void {
  const gate = byId("auth-gate");
  const tabs = byId("main-tabs");
  const panels = document.querySelector<HTMLElement>(".tab-panels");
  const settingsPanel = byId("settings-panel");
  const settingsToggle = byId("settings-toggle");
  const signOutBtn = byId<HTMLButtonElement>("auth-signout");
  const userLabel = byId("auth-user");
  const status = byId("status-strip");

  if (state) {
    gate.hidden = true;
    tabs.hidden = false;
    if (panels) panels.hidden = false;
    settingsToggle.hidden = false;
    signOutBtn.hidden = false;
    status.hidden = false;
    if (state.email) {
      userLabel.textContent = state.email;
      userLabel.hidden = false;
    } else {
      userLabel.textContent = "";
      userLabel.hidden = true;
    }
  } else {
    gate.hidden = false;
    tabs.hidden = true;
    if (panels) panels.hidden = true;
    settingsPanel.hidden = true;
    settingsToggle.hidden = true;
    signOutBtn.hidden = true;
    userLabel.hidden = true;
    userLabel.textContent = "";
    status.hidden = true;
  }
}

function bootstrapOnce(): void {
  if (tabsBootstrapped) return;
  tabsBootstrapped = true;
  try {
    bootstrap();
  } catch (err) {
    tabsBootstrapped = false;
    console.error("Finalysis bootstrap failed", err);
    setStatus(err instanceof Error ? err.message : "Failed to start the add-in.", "error");
  }
}

function bootstrap(): void {
  initTabs();
  initSettingsPanel();

  initDashboardTab();
  initDataImportTab();
  initReconciliationTab();
  initPortfolioTab();

  clearStatus();
}

function showSideloadMessage(): void {
  const msg = document.getElementById("sideload-msg");
  if (msg) msg.hidden = false;
}

/* --- Tab router --- */

function initTabs(): void {
  const buttons = document.querySelectorAll<HTMLButtonElement>(".tabs__btn");
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab || ""));
  });
}

function activateTab(tabId: string): void {
  if (!tabId) return;
  document.querySelectorAll<HTMLButtonElement>(".tabs__btn").forEach((b) => {
    b.classList.toggle("is-active", b.dataset.tab === tabId);
  });
  document.querySelectorAll<HTMLElement>(".tab-panel").forEach((panel) => {
    const active = panel.id === `tab-${tabId}`;
    panel.classList.toggle("is-active", active);
    panel.hidden = !active;
  });
}

/* --- Settings panel --- */

function initSettingsPanel(): void {
  const panel = byId("settings-panel");
  const toggle = byId<HTMLButtonElement>("settings-toggle");
  const input = byId<HTMLInputElement>("settings-api-key");
  const status = byId("settings-key-status");
  const saveBtn = byId<HTMLButtonElement>("settings-save");
  const clearBtn = byId<HTMLButtonElement>("settings-clear");

  const refreshStatus = () => {
    const key = loadApiKey();
    if (key) {
      status.textContent = `Saved key: ${maskApiKey(key)}`;
    } else {
      status.textContent = "No key saved.";
    }
  };

  toggle.addEventListener("click", () => {
    panel.hidden = !panel.hidden;
    if (!panel.hidden) {
      input.value = "";
      input.focus();
      refreshStatus();
    }
  });

  saveBtn.addEventListener("click", () => {
    void (async () => {
      const key = input.value.trim();
      if (!key) {
        setStatus("Enter an API key before saving.", "error");
        return;
      }
      try {
        await saveApiKey(key);
        input.value = "";
        refreshStatus();
        setStatus("API key saved.", "success");
      } catch (err) {
        setStatus(err instanceof Error ? err.message : String(err), "error");
      }
    })();
  });

  clearBtn.addEventListener("click", () => {
    void (async () => {
      try {
        await saveApiKey("");
        input.value = "";
        refreshStatus();
        setStatus("API key cleared.", "info");
      } catch (err) {
        setStatus(err instanceof Error ? err.message : String(err), "error");
      }
    })();
  });

  // Open settings automatically on first run if no key is saved.
  const existing = loadApiKey();
  if (!existing) {
    panel.hidden = false;
    status.textContent = "No key saved. Paste your Finnhub API key to get started.";
  } else {
    refreshStatus();
  }
}
