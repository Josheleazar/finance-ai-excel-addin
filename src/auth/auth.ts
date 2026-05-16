/*
 * Auth dialog entry point.
 *
 * This file runs INSIDE an Office Dialog window opened by the task pane via
 * `Office.context.ui.displayDialogAsync`. Because the dialog is a top-level
 * browser window (not an iframe), Clerk's cookies set here are first-party
 * and the sign-in flow works on Excel for the web, Mac, and Windows.
 *
 * Why we load Clerk from the CDN instead of bundling it:
 *   The npm package `@clerk/clerk-js` v6 ships a lightweight loader
 *   (`clerk.js` / `clerk.mjs`) that lazy-imports its UI component chunks at
 *   runtime via webpack-style dynamic imports. Inside the Office Dialog
 *   webview those chunk URLs don't resolve, so calling `mountSignIn()`
 *   throws "Clerk was not loaded with UI components". Loading the all-in-one
 *   `clerk.browser.js` build from Clerk's own CDN sidesteps the chunk
 *   problem entirely — it's the officially-supported pattern for vanilla JS.
 *
 *   We still depend on `@clerk/clerk-js` so we get accurate TypeScript types
 *   (`import type` is erased at compile time and adds nothing to the bundle).
 *
 * Flow:
 *   1. Office.onReady — confirm the dialog API is available.
 *   2. Inject Clerk's CDN script using the FAPI host decoded from the
 *      publishable key. The `clerk.browser.js` bundle auto-instantiates
 *      `window.Clerk` (using the `data-clerk-publishable-key` attribute on
 *      the script tag) — we then `await window.Clerk.load()` to initialize.
 *   3. If the user is already signed in (cookie survived), mint a JWT and
 *      messageParent immediately.
 *   4. Otherwise mountSignIn() and listen for sign-in completion.
 *   5. On success, mint a JWT via session.getToken({ template: 'office-addin' })
 *      and messageParent({ ok: true, token, ... }).
 *   6. On any failure, messageParent({ ok: false, error }).
 *
 * The task-pane side parses the JSON message in services/auth.ts.
 */

/* global Office, document, console, process, HTMLDivElement, window, atob */

import type { Clerk } from "@clerk/clerk-js";

declare const process: { env: { CLERK_PUBLISHABLE_KEY?: string } };

const PUBLISHABLE_KEY = process.env.CLERK_PUBLISHABLE_KEY || "";
const JWT_TEMPLATE = "office-addin";
// Major-version pin matches package.json. Clerk's CDN serves
// `/npm/@clerk/clerk-js@<major>/dist/clerk.browser.js`.
const CLERK_JS_MAJOR_VERSION = "6";

/**
 * Some Office Dialog hosts (notably the Trident/Edge-Legacy webview used by
 * Excel desktop on older Windows builds) do not expose `history.pushState` /
 * `history.replaceState`. Clerk calls these during its sign-in flow and
 * throws `globalThis.history.replaceState is not a function`. We install
 * harmless no-ops *before* loading Clerk so the call sites succeed. The
 * sign-in component is mounted with `routing: "virtual"` below, so we don't
 * actually need URL changes anyway.
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
 * Decode the Clerk Frontend API host from a publishable key.
 *
 * Publishable keys have the format `pk_(test|live)_<base64(fapi + "$")>`.
 * For example, `pk_test_Y2xlcmsuZXhhbXBsZS5jb20k` decodes to
 * `clerk.example.com$`, so the FAPI host is `clerk.example.com`.
 */
function fapiFromPublishableKey(key: string): string {
  const match = /^pk_(test|live)_(.+)$/.exec(key);
  if (!match) {
    throw new Error("CLERK_PUBLISHABLE_KEY is not in the expected pk_(test|live)_… format.");
  }
  let decoded: string;
  try {
    decoded = atob(match[2]);
  } catch {
    throw new Error("CLERK_PUBLISHABLE_KEY is not a valid base64-encoded publishable key.");
  }
  const fapi = decoded.endsWith("$") ? decoded.slice(0, -1) : decoded;
  if (!fapi) {
    throw new Error("CLERK_PUBLISHABLE_KEY does not contain a Frontend API host.");
  }
  return fapi;
}

/**
 * Inject Clerk's CDN script tag and resolve with a ready-to-load `Clerk`
 * instance. The `clerk.browser.js` bundle reads its publishable key from the
 * `data-clerk-publishable-key` attribute on the script element and ends with
 *
 *   window.Clerk || (window.Clerk = new Clerk(publishableKey, {...}))
 *
 * so by the time the script's `onload` fires, `window.Clerk` is already an
 * instance. We just validate the shape (must have a `.load()` method) and
 * return it.
 *
 * Loading is idempotent — if `window.Clerk` already exists from a previous
 * call we reuse it instead of re-injecting the script.
 */
async function loadClerkFromCdn(publishableKey: string): Promise<Clerk> {
  const readWindowClerk = (): Clerk | undefined => {
    const candidate = (window as unknown as { Clerk?: unknown }).Clerk;
    // A loaded Clerk instance is an object with a `.load` method.
    if (
      candidate &&
      typeof candidate === "object" &&
      typeof (candidate as { load?: unknown }).load === "function"
    ) {
      return candidate as Clerk;
    }
    return undefined;
  };

  // Already loaded? (e.g. dialog reload after partial flow.)
  const existing = readWindowClerk();
  if (existing) return existing;

  const fapi = fapiFromPublishableKey(publishableKey);
  const scriptUrl = `https://${fapi}/npm/@clerk/clerk-js@${CLERK_JS_MAJOR_VERSION}/dist/clerk.browser.js`;

  await new Promise<void>((resolve, reject) => {
    const script = document.createElement("script");
    script.src = scriptUrl;
    script.async = true;
    script.crossOrigin = "anonymous";
    // The bundle reads this attribute to auto-instantiate `window.Clerk`
    // with the correct publishable key.
    script.setAttribute("data-clerk-publishable-key", publishableKey);
    script.onload = () => resolve();
    script.onerror = () =>
      reject(new Error(`Failed to load Clerk from ${scriptUrl}. Check network and AppDomains.`));
    document.head.appendChild(script);
  });

  const instance = readWindowClerk();
  if (!instance) {
    throw new Error(
      "Clerk script loaded but window.Clerk is not a usable instance. " +
        "This usually means the publishable key on the script tag was not picked up."
    );
  }
  return instance;
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
    clerk = await loadClerkFromCdn(PUBLISHABLE_KEY);
  } catch (err) {
    reportError(err instanceof Error ? err.message : "Failed to load Clerk.");
    return;
  }

  try {
    // The bundle auto-instantiates Clerk with the publishable key from the
    // script-tag attribute, but we still have to call `.load()` to fetch the
    // session/environment from the Frontend API.
    await clerk.load();
  } catch (err) {
    reportError(err instanceof Error ? err.message : "Failed to initialize Clerk.");
    return;
  }

  // Already signed in (e.g. session cookie persisted in this dialog window).
  if (clerk.user && clerk.session) {
    const posted = await mintAndPostToken(clerk);
    if (posted) return;
  }

  // Mount the sign-in widget and wait for the user to complete it.
  setStatus("Sign in to continue.");
  const mount = document.getElementById("clerk-mount") as HTMLDivElement | null;
  if (!mount) {
    reportError("Auth dialog DOM is missing #clerk-mount.");
    return;
  }
  mount.innerHTML = ""; // clear any fallback content

  try {
    // `routing: "virtual"` keeps Clerk's flow entirely in-memory and avoids
    // any calls into `window.history.{push,replace}State`, which the Office
    // Dialog webview may not implement. The addListener() callback below
    // detects sign-in completion and mints the JWT — we don't need a
    // redirect URL at all in virtual mode.
    //
    // Note: @clerk/clerk-js's vanilla SignInProps types currently advertise
    // only "path" | "hash", but the runtime accepts "virtual" (it's the same
    // mode the React SDK exposes for embedded contexts). Cast through unknown
    // to satisfy the stale type while keeping the rest of the props checked.
    clerk.mountSignIn(mount, {
      routing: "virtual",
    } as unknown as Parameters<Clerk["mountSignIn"]>[1]);
  } catch (err) {
    reportError(err instanceof Error ? err.message : "Failed to render Clerk sign-in.");
    return;
  }

  // React to sign-in / session changes. Only post once so the dialog can close
  // cleanly without racing duplicate handlers.
  let posted = false;
  clerk.addListener(async () => {
    if (posted) return;
    if (clerk.user && clerk.session) {
      posted = true;
      const ok = await mintAndPostToken(clerk);
      if (!ok) posted = false; // allow retry if the mint failed
    }
  });
}

Office.onReady(() => {
  void main().catch((err) => {
    reportError(err instanceof Error ? err.message : String(err));
  });
});
