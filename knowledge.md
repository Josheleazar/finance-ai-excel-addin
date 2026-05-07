# Project knowledge

## What this is
- **finance-ai-excel-addin** — an Office Add-in (Excel task pane) scaffolded from Microsoft's `Office-Addin-TaskPane` TypeScript template.
- The actual add-in code lives in the `finalysis/` subdirectory (not the repo root). Treat `finalysis/` as the working project root for install/build/dev.
- The add-in is named **"finalysis"** in `manifest.xml`. It targets Excel (`<Hosts><Host Name="Workbook"/></Hosts>`) and currently ships the stock template behavior (a "Run" button that highlights the selected range yellow).

## Key directories and files
- `finalysis/src/taskpane/` — task pane UI
  - `taskpane.ts` — entry point; registers `Office.onReady` and the `run()` handler that uses the Excel JS API.
  - `taskpane.html` — pane markup (Fluent UI / Fabric CSS).
  - `taskpane.css` — pane styles.
- `finalysis/src/commands/` — ribbon command function file
  - `commands.ts` — `action()` function registered via `Office.actions.associate`. NOTE: the template code calls `Office.context.mailbox.item.notificationMessages` which is **Outlook-only** and will throw in Excel. Replace before wiring up a real command.
  - `commands.html` — host page for the function file.
- `finalysis/manifest.xml` — Office Add-in manifest (IDs, URLs, ribbon buttons). Dev URL is `https://localhost:3000/`. Prod URL placeholder is `https://www.contoso.com/` in `webpack.config.js` — change before shipping.
- `finalysis/webpack.config.js` — webpack config with three entries: `polyfill`, `taskpane`, `commands`. Serves HTTPS on port 3000 using dev certs from `office-addin-dev-certs`.
- `finalysis/dist/` — built output (checked-in currently; typically a build artifact).
- `finalysis/tsconfig.json` — `target: es5`, `jsx: react`, `allowJs: true`.
- `finalysis/babel.config.json` — `@babel/preset-env` + `@babel/preset-typescript` (TS is transpiled by Babel, not `tsc`).
- `.agents/types/` — Codebuff agent type definitions; not part of the add-in build.

## Commands (run from `finalysis/`)
- Install: `npm install`
- Dev server (HTTPS on :3000): `npm run dev-server`
- Start Excel with the add-in sideloaded: `npm start`
- Stop the debugging session: `npm stop`
- Production build: `npm run build` (outputs to `dist/`)
- Dev build: `npm run build:dev`
- Watch build: `npm run watch`
- Lint / fix / format: `npm run lint` / `npm run lint:fix` / `npm run prettier`
- Validate manifest: `npm run validate`
- M365 sign-in/out (for dev): `npm run signin` / `npm run signout`

## Conventions and gotchas
- **Work inside `finalysis/`** for all npm/webpack commands. Running them from the repo root will fail (no `package.json` there).
- **No test framework is configured.** There is no `npm test` script. Add Jest (or similar) before writing tests.
- **TypeScript is compiled via Babel**, so `tsc` is not used for emission — type errors do not block builds. Run the editor's TS server or add a `tsc --noEmit` step if you want type checking in CI.
- **Target is ES5 + IE11** (`browserslist: ["last 2 versions", "ie 11"]`, `tsconfig target: es5`). Avoid modern-only APIs or add polyfills via the `polyfill` entry.
- **HTTPS is required.** The Office client loads the pane over HTTPS. `office-addin-dev-certs` installs a local CA; first run may prompt for trust.
- **Port 3000 is hardcoded** in `manifest.xml` (`https://localhost:3000/...`). Changing the port requires updating the manifest too.
- **Production URL placeholder**: `webpack.config.js` rewrites `https://localhost:3000/` → `https://www.contoso.com/` in the manifest during prod builds. Update `urlProd` before deploying.
- **`commands.ts` is Outlook-template code.** The `mailbox.item.notificationMessages` API does not exist in Excel. Replace with Excel-appropriate logic (e.g., a direct `Excel.run` call) before exposing the ribbon command.
- **Linting** uses `eslint-plugin-office-addins` (recommended config). Prettier config is `office-addin-prettier-config`.
- **The repo has a two-level layout** (root + `finalysis/`). If you add CI or tooling at the root, remember to `cd finalysis` first.

## Useful links
- Office.js reference: https://learn.microsoft.com/javascript/api/overview
- Excel JS API: https://learn.microsoft.com/javascript/api/excel
- Sideloading an add-in: https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins

## Progress log

### Pass 1 — Foundation + Data Import (complete)
Scope agreed with user: restructure, 4-tab shell, wire all 4 Finnhub endpoints, Office document-settings API key w/ localStorage fallback, direct API calls (proxy TODO), leave manifest alone.

**New structure under `finalysis/src/taskpane/`:**
- `types/finnhub.ts` — `Quote`, `CandleResponse`, `CompanyProfile`, `BasicFinancials`.
- `services/finnhub.ts` — `getQuote`, `getCandles`, `getCompanyProfile`, `getBasicFinancials`. Exports `FINNHUB_BASE_URL`. Has TODO for backend proxy. Friendly messages for 401 (bad key) / 403 (candles premium-gated) / 429 (rate limit).
- `services/settings.ts` — `loadApiKey` / `saveApiKey` / `clearApiKey` / `maskApiKey`. Uses `Office.context.document.settings` + `saveAsync`; falls back to `localStorage` key `finalysis.finnhubApiKey`.
- `services/excel.ts` — batched `writeTableToNewSheet` / `writeKeyValueToNewSheet` (single `Excel.run`, auto-named sheets, autofit columns).
- `services/ui.ts` — `setStatus`, `byId`, `inputValue`, `formatUnix`.
- `tabs/dashboard.ts` — in-pane quote lookup card (no sheet write).
- `tabs/dataImport.ts` — 4 Finnhub imports, each writes to a new sheet.
- `tabs/reconciliation.ts`, `tabs/aiInsights.ts` — init stubs; markup is static placeholders in `taskpane.html`.

**Rewrites:**
- `taskpane.html` — header w/ settings gear, 4-tab nav, settings panel, per-action cards, status strip. Removed Fabric CDN.
- `taskpane.css` — dependency-free Fluent-inspired styles. Includes `.card__body--note` for the candles premium hint.
- `taskpane.ts` — `Office.onReady` + Excel host check, tab router, settings panel (auto-opens on first run if no key).
- `commands/commands.ts` — removed Outlook-only `notificationMessages` call; safe no-op.

**Tooling:**
- Added `typecheck` npm script (`tsc --noEmit`) to `finalysis/package.json` (Babel doesn't type-check).

**Validation at end of pass:** `tsc --noEmit` ✅ 0 errors. `npm run lint` ✅ 0 errors / 0 warnings. Code-reviewer ✅ no blocking issues.

### Known / deferred for next pass
- **Reconciliation tab** — UI is a static placeholder; `services/reconcile.ts` pure-matching module not yet written. Plan: two-column matcher (select two ranges, output new `Reconciliation_<ts>` sheet with matches / A-only / B-only / duplicates, optional numeric tolerance).
- **AI Insights tab** — placeholder only; no provider wired.
- **Finnhub `/stock/candle` is premium** on current Finnhub pricing. Free-tier users will get the 403 message. Consider swapping to a free-tier alternative (e.g., `/stock/symbol` + another source) or gating the card.
- **Backend proxy** — API key still ships to the browser. `FINNHUB_BASE_URL` is a single constant with a TODO; swap to our proxy origin when built.
- **Manifest metadata** — still says `DisplayName="finalysis"`, `ProviderName="Contoso"`, `Description="A template to get started."`, prod URL `contoso.com`. Update before any non-dev distribution.
- **Icons** — `finalysis/assets/` exists on disk (the initial file-tree snapshot was stale); manifest + webpack `CopyWebpackPlugin` references are valid.
- **Tests** — still no test framework. `services/reconcile.ts` (future) and `services/finnhub.ts` (URL building, error mapping) are the obvious first unit-test targets; Jest is the natural pick.
- **CI** — nothing yet. Minimum useful job: `cd finalysis && npm ci && npm run typecheck && npm run lint && npm run build`.
