/*
 * Auth dialog entry point.
 *
 * This file runs INSIDE an Office Dialog window opened by the task pane via
 * `Office.context.ui.displayDialogAsync`. Because the dialog is a top-level
 * browser window (not an iframe), Clerk's cookies set here are first-party
 * and the sign-in flow works on Excel for the web, Mac, and Windows.
 *
 * Flow:
 *   1. Office.onReady — confirm the dialog API is available.
 *   2. `new Clerk(publishableKey)` + `await clerk.load()`.
 *   3. If the user is already signed in (cookie survived), mint a JWT and
 *      messageParent immediately.
 *   4. Otherwise mountSignIn() with `routing: "hash"` (URL fragments only;
 *      no history API calls) and listen for sign-in completion.
 *   5. On success, mint a JWT via session.getToken({ template: 'office-addin' })
 *      and messageParent({ ok: true, token, ... }).
 *   6. On any failure, messageParent({ ok: false, error }).
 *
 * The task-pane side parses the JSON message in services/auth.ts.
 *
 * Why hash routing:
 *   The vanilla `@clerk/clerk-js` SDK only supports `routing: "path" | "hash"`
 *   (the `"virtual"` mode exists only in `@clerk/clerk-react`). Path routing
 *   calls `window.history.replaceState`, which the Office Dialog webview can
 *   throw on. Hash routing uses URL fragments and never touches history, so
 *   it's the right fit for embedded dialog contexts.
 *
 * Why the default `@clerk/clerk-js` bundle (NOT /no-rhc):
 *   Clerk ships two builds. The default bundle (`@clerk/clerk-js`) fetches its
 *   UI implementation from Clerk's CDN at runtime; it auto-wires the internal
 *   `ClerkUI` controller during `clerk.load()`, so `mountSignIn` / `openSignUp`
 *   work out of the box. The `/no-rhc` ("no remote hosted components") subpath
 *   is a headless-only build: it does NOT auto-wire UI components, and the
 *   `ClerkUI` constructor it expects via `new Clerk(key, { ui: { ClerkUI } })`
 *   is not exported by the package — `/no-rhc` is intended for callers that
 *   render their own UI on top of Clerk's API (`clerk.client.signIn.create`).
 *   Using `/no-rhc` with `mountSignIn` produces the misleading
 *   `Error("Clerk was not loaded with Ui components")` thrown by
 *   `assertComponentsReady`. The Office Dialog opened via `displayDialogAsync`
 *   is a real top-level browser window with normal network access, so the
 *   CDN fetch the default bundle performs is fine. The Clerk CDN hosts
 *   (*.clerk.accounts.dev, *.accounts.dev, img.clerk.com, challenges.cloudflare.com)
 *   are whitelisted in manifest.xml's <AppDomains>.
 */

/* global Office, document, console, process, HTMLDivElement, window */

// `Clerk` is a NAMED export from @clerk/clerk-js (no default export). Using
// `import Clerk from ...` under Babel's CJS interop resolves to `module.default`,
// which is undefined — surfacing as `TypeError: a.default is not a constructor`
// when we later call `new Clerk(...)`. Always use the named import.
import { Clerk } from "@clerk/clerk-js";

declare const process: { env: { CLERK_PUBLISHABLE_KEY?: string } };

const PUBLISHABLE_KEY = process.env.CLERK_PUBLISHABLE_KEY || "";
const JWT_TEMPLATE = "office-addin";

/**
 * Some Office Dialog hosts (notably the Trident/Edge-Legacy webview used by
 * Excel desktop on older Windows builds) do not expose `history.pushState` /
 * `history.replaceState`. We use `routing: "hash"` below so Clerk should
 * never call those, but installing harmless no-ops as a defensive polyfill
 * costs nothing and protects against any other code path that might touch
 * the history API.
 */
function patchHistoryApi(): void {
  try {
    if (typeof window === "undefined" || !window.history) return;
    const h = window.history as History & {
      pushState?: typeof History.prototype.pushState;
      replaceState?: typeof History.prototype.replaceState;
    };
    if (typeof h.pushState !== "function") {
      h.pushState = function () {
        /* no-op: Office Dialog host lacks history API */
      } as typeof History.prototype.pushState;
    }
    if (typeof h.replaceState !== "function") {
      h.replaceState = function () {
        /* no-op: Office Dialog host lacks history API */
      } as typeof History.prototype.replaceState;
    }
  } catch {
    /* ignore — patching is best-effort */
  }
}
patchHistoryApi();

interface AuthSuccess {
  ok: true;
  token: string;
  expiresAt: number; // ms epoch
  userId: string;
  email: string | null;
}

interface AuthFailure {
  ok: false;
  error: string;
}

type AuthMessage = AuthSuccess | AuthFailure;

function setStatus(message: string, kind: "info" | "error" = "info"): void {
  const el = document.getElementById("auth-status");
  if (!el) return;
  el.textContent = message;
  el.dataset.kind = kind;
}

function showFallback(html: string): void {
  const mount = document.getElementById("clerk-mount");
  if (!mount) return;
  mount.innerHTML = `<div class="auth-fallback">${html}</div>`;
}

function postToParent(message: AuthMessage): void {
  try {
    if (
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.ui &&
      typeof Office.context.ui.messageParent === "function"
    ) {
      Office.context.ui.messageParent(JSON.stringify(message));
    } else {
      // Page opened outside an Office dialog (e.g. dev preview). Surface the
      // payload so a developer can inspect it.
      // eslint-disable-next-line no-console
      console.warn("[finalysis auth] messageParent unavailable; payload:", message);
    }
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error("[finalysis auth] messageParent failed", err);
  }
}

function reportError(error: string): void {
  setStatus(error, "error");
  postToParent({ ok: false, error });
}

/**
 * Pull a fresh JWT for the current Clerk session and ship it to the parent.
 * Returns true if a token was successfully posted.
 */
async function mintAndPostToken(clerk: Clerk): Promise<boolean> {
  const session = clerk.session;
  const user = clerk.user;
  if (!session || !user) return false;

  let token: string | null = null;
  try {
    token = await session.getToken({ template: JWT_TEMPLATE });
  } catch (err) {
    // If the JWT template doesn't exist yet, fall back to the default token so
    // we can at least confirm sign-in succeeded.
    // eslint-disable-next-line no-console
    console.warn(
      `[finalysis auth] getToken({ template: '${JWT_TEMPLATE}' }) failed; ` +
        "falling back to default session token. Create the template in the " +
        "Clerk dashboard for production. Reason:",
      err
    );
    try {
      token = await session.getToken();
    } catch (err2) {
      reportError(err2 instanceof Error ? err2.message : "Failed to mint session token.");
      return false;
    }
  }
  if (!token) {
    reportError("Clerk did not return a session token.");
    return false;
  }

  // Clerk JWTs default to a 60s lifetime. Store an absolute expiry so the
  // task pane can decide when to re-mint. We deliberately under-estimate by
  // 5 seconds to leave room for clock skew.
  const expiresAt = Date.now() + 55_000;
  const email = user.primaryEmailAddress?.emailAddress ?? null;

  postToParent({ ok: true, token, expiresAt, userId: user.id, email });
  setStatus("Signed in. You can close this window.", "info");
  return true;
}

async function main(): Promise<void> {
  if (!PUBLISHABLE_KEY) {
    setStatus("Configuration error.", "error");
    showFallback(
      "<h2>Missing Clerk publishable key</h2>" +
        "<p>Set <code>CLERK_PUBLISHABLE_KEY</code> in your <code>.env</code> " +
        "(local) or Vercel environment (deploy) and rebuild.</p>"
    );
    postToParent({ ok: false, error: "CLERK_PUBLISHABLE_KEY is not configured." });
    return;
  }

  setStatus("Loading Clerk…");

  let clerk: Clerk;
  try {
    clerk = new Clerk(PUBLISHABLE_KEY);
    await clerk.load();

    // eslint-disable-next-line no-console
    console.log("[finalysis auth] Clerk loaded", {
      hasMountSignIn: typeof clerk.mountSignIn === "function",
      hasOpenSignIn: typeof clerk.openSignIn === "function",
    });
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error("[finalysis auth] Clerk load failed:", err);
    reportError(err instanceof Error ? err.message : String(err));
    return;
  }

  // Already signed in (session cookie still valid)
  if (clerk.user && clerk.session) {
    const posted = await mintAndPostToken(clerk);
    if (posted) return;
  }

  // Mount the sign-in form
  setStatus("Sign in to continue.");
  const mount = document.getElementById("clerk-mount") as HTMLDivElement | null;
  if (!mount) {
    reportError("Auth dialog DOM is missing #clerk-mount.");
    return;
  }
  mount.innerHTML = "";

  try {
    console.log("🚀 Mounting Clerk sign-in UI...");
    clerk.mountSignIn(mount, {
      routing: "hash",
    });
  } catch (err) {
    console.error("❌ mountSignIn failed:", err);
    reportError(err instanceof Error ? err.message : "Failed to render sign-in form.");
  }

  // Listen for successful sign-in
  let posted = false;
  clerk.addListener(async () => {
    if (posted) return;
    if (clerk.user && clerk.session) {
      posted = true;
      await mintAndPostToken(clerk);
    }
  });
}

Office.onReady(() => {
  void main().catch((err) => {
    reportError(err instanceof Error ? err.message : String(err));
  });
});
