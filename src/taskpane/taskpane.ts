/*
 * Finalysis task pane bootstrap.
 *
 * Responsibilities:
 *   - Wait for Office.onReady and verify the host is Excel.
 *   - Hydrate the settings panel (API key) from storage.
 *   - Wire up the tab switcher.
 *   - Delegate per-tab event handlers to their modules.
 */

/* global console, document, Office, HTMLElement, HTMLButtonElement, HTMLInputElement */

import { initDashboardTab } from "./tabs/dashboard";
import { initDataImportTab } from "./tabs/dataImport";
import { initReconciliationTab } from "./tabs/reconciliation";
import { initPortfolioTab } from "./tabs/portfolio";
import { loadApiKey, maskApiKey, saveApiKey } from "./services/settings";
import { byId, clearStatus, setStatus } from "./services/ui";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    showSideloadMessage();
    return;
  }

  try {
    bootstrap();
  } catch (err) {
    console.error("Finalysis bootstrap failed", err);
    setStatus(err instanceof Error ? err.message : "Failed to start the add-in.", "error");
  }
});

function bootstrap(): void {
  byId("app-root").classList.remove("is-loading");

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
