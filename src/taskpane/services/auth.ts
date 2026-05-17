/*
 * Task-pane-side authentication service.
 *
 * Responsibilities:
 *   - Open the Clerk sign-in dialog via Office.context.ui.displayDialogAsync
 *     (the only window primitive that works reliably inside an Excel task
 *     pane on desktop, Mac, and the web).
 *   - Persist the resulting JWT + user info in OfficeRuntime.storage
 *     (encrypted, add-in-scoped). Falls back to localStorage when running
 *     outside Office (e.g. browser preview).
 *   - Provide getFreshToken(): returns the cached JWT if it has comfortable
 *     headroom, otherwise re-opens the dialog to mint a new one.
 *   - signOut(): clears local state and asks Clerk (in a dialog) to end the
 *     session.
 *
 * Notes:
 *   - We do NOT load @clerk/clerk-js inside the task pane. It would attempt
 *     to use third-party cookies (blocked in Excel for the web / Safari) and
 *     would also bloat the taskpane bundle. All Clerk interaction happens
 *     in the auth.html dialog.
 *   - Subscribers (UI code) can listen for auth-state changes via
 *     onAuthStateChange().
 */

/* global Office, OfficeRuntime, window, console, URL */

export interface AuthState {
  token: string;
  /** ms epoch when `token` should be considered expired */
  expiresAt: number;
  userId: string;
  email: string | null;
}

interface DialogSuccessPayload {
  ok: true;
  token: string;
  expiresAt: number;
  userId: string;
  email: string | null;
}

interface DialogFailurePayload {
  ok: false;
  error: string;
}

type DialogPayload = DialogSuccessPayload | DialogFailurePayload;

const STORAGE_KEY = "finalysis.auth.v1";
/**
 * If a cached token has fewer than this many milliseconds of life left, treat
 * it as expired and re-mint. Clerk JWTs default to 60s; 10s of headroom is
 * generous without forcing a dialog on every API call.
 */
const TOKEN_REFRESH_MARGIN_MS = 10_000;

type AuthListener = (state: AuthState | null) => void;
const listeners = new Set<AuthListener>();

let cached: AuthState | null = null;
let hydrated = false;

/* -------------------- Storage adapter -------------------- */

function hasOfficeRuntimeStorage(): boolean {
  try {
    return (
      typeof OfficeRuntime !== "undefined" &&
      !!OfficeRuntime &&
      !!OfficeRuntime.storage &&
      typeof OfficeRuntime.storage.getItem === "function"
    );
  } catch {
    return false;
  }
}

async function storageGet(): Promise<string | null> {
  if (hasOfficeRuntimeStorage()) {
    try {
      const v = await OfficeRuntime.storage.getItem(STORAGE_KEY);
      return v ?? null;
    } catch {
      // fall through to localStorage
    }
  }
  try {
    return window.localStorage.getItem(STORAGE_KEY);
  } catch {
    return null;
  }
}

async function storageSet(value: string | null): Promise<void> {
  if (hasOfficeRuntimeStorage()) {
    try {
      if (value === null) {
        await OfficeRuntime.storage.removeItem(STORAGE_KEY);
      } else {
        await OfficeRuntime.storage.setItem(STORAGE_KEY, value);
      }
      return;
    } catch {
      // fall through to localStorage
    }
  }
  try {
    if (value === null) {
      window.localStorage.removeItem(STORAGE_KEY);
    } else {
      window.localStorage.setItem(STORAGE_KEY, value);
    }
  } catch {
    // localStorage may be disabled — best effort.
  }
}

/* -------------------- State -------------------- */

function notify(): void {
  listeners.forEach((fn) => {
    try {
      fn(cached);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error("[finalysis auth] listener threw", err);
    }
  });
}

async function hydrate(): Promise<void> {
  if (hydrated) return;
  hydrated = true;
  const raw = await storageGet();
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw) as Partial<AuthState>;
    if (
      parsed &&
      typeof parsed.token === "string" &&
      typeof parsed.expiresAt === "number" &&
      typeof parsed.userId === "string"
    ) {
      cached = {
        token: parsed.token,
        expiresAt: parsed.expiresAt,
        userId: parsed.userId,
        email: typeof parsed.email === "string" ? parsed.email : null,
      };
    }
  } catch {
    // Corrupt entry — wipe it.
    await storageSet(null);
  }
}

export async function getStoredAuth(): Promise<AuthState | null> {
  await hydrate();
  return cached;
}

export async function setStoredAuth(state: AuthState | null): Promise<void> {
  await hydrate();
  cached = state;
  await storageSet(state ? JSON.stringify(state) : null);
  notify();
}

export async function isAuthenticated(): Promise<boolean> {
  const s = await getStoredAuth();
  return !!s;
}

export function onAuthStateChange(listener: AuthListener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/* -------------------- Dialog flow -------------------- */

function authUrl(): string {
  // The dialog must be served from the same origin as the task pane so that
  // Office can bridge messageParent. webpack emits auth.html alongside
  // taskpane.html, so resolving relative to the current location works in
  // both dev (https://localhost:3000) and prod.
  return new URL("auth.html", window.location.href).toString();
}

/**
 * Open the auth dialog and resolve once it posts a payload back (or the user
 * closes it).
 */
function openDialog(options: { silent?: boolean } = {}): Promise<{ payload: DialogPayload }> {
  return new Promise((resolve, reject) => {
    if (
      !Office ||
      !Office.context ||
      !Office.context.ui ||
      typeof Office.context.ui.displayDialogAsync !== "function"
    ) {
      reject(new Error("Office Dialog API is not available in this host."));
      return;
    }

    const { silent = false } = options;
    Office.context.ui.displayDialogAsync(
      authUrl(),
      {
        height: silent ? 30 : 65,
        width: silent ? 30 : 35,
        promptBeforeOpen: false,
      },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(asyncResult.error?.message || "Failed to open sign-in dialog."));
          return;
        }
        const dialog = asyncResult.value;
        let resolved = false;

        const finish = (payload: DialogPayload) => {
          if (resolved) return;
          resolved = true;

          // ==================== DEBUG MODE - Keep dialog open ====================
          console.log("[DEBUG] Auth payload received:", payload);
          
          // Keep the dialog open for 10 seconds so you can inspect it
          setTimeout(() => {
            try {
              dialog.close();
            } catch (e) {
              // already closed
            }
            resolve({ payload });
          }, 10000); // 10 seconds - plenty of time to open DevTools
          // =====================================================================
        };

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          // The dialog posts JSON via Office.context.ui.messageParent.
          const message = (arg as { message?: string }).message ?? "";
          let payload: DialogPayload;
          try {
            payload = JSON.parse(message) as DialogPayload;
          } catch {
            payload = { ok: false, error: "Malformed message from auth dialog." };
          }
          finish(payload);
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          // 12006 = user closed the dialog manually. Other codes are network
          // / unloaded errors. Either way we surface a cancellation.
          const code = (arg as { error?: number }).error ?? 0;
          finish({
            ok: false,
            error:
              code === 12006
                ? "Sign-in was cancelled."
                : `Sign-in dialog closed unexpectedly (code ${code}).`,
          });
        });
      }
    );
  });
}

/**
 * Open the dialog for an interactive sign-in. Resolves with the new auth
 * state on success; rejects on failure.
 */
export async function openSignInDialog(): Promise<AuthState> {
  return runAuthDialog({ silent: false });
}

async function runAuthDialog(options: { silent: boolean }): Promise<AuthState> {
  const { payload } = await openDialog(options);
  if (payload.ok !== true) {
    // tsconfig has strictNullChecks off, so the discriminated union doesn't
    // auto-narrow here; explicit cast keeps both flavors of tsc happy.
    throw new Error((payload as DialogFailurePayload).error);
  }
  const success = payload as DialogSuccessPayload;
  const state: AuthState = {
    token: success.token,
    expiresAt: success.expiresAt,
    userId: success.userId,
    email: success.email,
  };
  await setStoredAuth(state);
  return state;
}

/**
 * Return a JWT that the caller can attach to a backend request. If the
 * cached token has comfortable headroom we return it directly; otherwise we
 * pop a tiny dialog so Clerk can mint a fresh one. The dialog is opened in
 * "silent" mode (a small window) because, if the Clerk session cookie is
 * still alive in the dialog window, auth.ts mints and posts immediately
 * without rendering the sign-in form.
 *
 * Throws if the user is not signed in or cancels the refresh.
 */
export async function getFreshToken(): Promise<string> {
  const current = await getStoredAuth();
  if (!current) {
    throw new Error("Not signed in.");
  }
  const headroom = current.expiresAt - Date.now();
  if (headroom > TOKEN_REFRESH_MARGIN_MS) {
    return current.token;
  }
  // Try a tiny background dialog first (Clerk session cookie should still be
  // alive in the dialog window, so auth.ts mints + posts immediately without
  // rendering the sign-in form). If that fails — e.g. the Clerk session has
  // expired and Clerk wants to re-prompt for credentials — fall back to a
  // full-size dialog so the user can actually see and use the form.
  try {
    const next = await runAuthDialog({ silent: true });
    return next.token;
  } catch (err) {
    // eslint-disable-next-line no-console
    console.warn("[finalysis auth] silent refresh failed, opening full dialog", err);
    const next = await runAuthDialog({ silent: false });
    return next.token;
  }
}

/**
 * Clear *local* auth state from this task pane.
 *
 * IMPORTANT: this only wipes the cached JWT in OfficeRuntime.storage. The
 * Clerk session cookie set inside the dialog window is still alive, so the
 * very next `openSignInDialog()` call will silently re-authenticate the
 * same user without prompting. That's the right behavior for an "end this
 * session in the task pane" button; if you need a true "switch accounts" /
 * "sign out everywhere" affordance, call `clerk.signOut()` from inside the
 * dialog (TODO: add a `?signout=1` query param + handler in auth.ts).
 */
export async function signOut(): Promise<void> {
  await setStoredAuth(null);
}
