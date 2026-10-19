# Probability Step Chart — Excel add-in

Turn a discrete probability distribution into the duplicated X/Y coordinates of its
**cumulative distribution function (CDF)**, draw it as an **XY scatter chart with straight
lines**, compare several distributions on the same axes, and test **first- and
second-order stochastic dominance** — directly inside Excel for Mac, Windows and the web.

| Input (selected range) | | → | Generated table | |
|---|---|---|---|---|
| **Outcome** | **Probability** | | **X (Outcome)** | **Y (Cumulative Probability)** |
| 11 | 0.4 | | 0 | 0 |
| 32 | 0.3 | | 11 | 0 |
| 53 | 0.3 | | 11 | 0.4 |
| | | | 32 | 0.4 |
| | | | 32 | 0.7 |
| | | | 53 | 0.7 |
| | | | 53 | 1 |

---

## Contents

- [Features](#features)
- [Requirements and compatibility](#requirements-and-compatibility)
- [Quick start (local development)](#quick-start-local-development)
- [Sideloading on Mac, Windows and the web](#sideloading)
- [Using the add-in](#using-the-add-in)
- [How the calculations work](#how-the-calculations-work)
- [Architecture](#architecture)
- [Testing](#testing)
- [Debugging](#debugging)
- [Packaging and distribution](#packaging-and-distribution)
- [Troubleshooting](#troubleshooting)
- [Known limitations](#known-limitations)

---

## Features

**Input**
- Reads the current selection (a single cell expands to its surrounding table; whole-column selections are trimmed to the data).
- Detects a header row automatically, with a manual override.
- Validates outcomes are numeric, probabilities are within [0, 1] and sum to 1 within a configurable tolerance (±1e−9 … ±1e−2).
- Sorts outcomes, merges duplicate outcomes (summing their probabilities), keeps zero-probability outcomes with a warning.
- Accepts decimals, cells formatted as Excel percentages, text such as `40%`, and percentage points (40/30/30, auto-detected or forced).
- Shows every problem with its cell address, never modifies the source data.
- Uses the text above the range (e.g. “Make Component”) as the default distribution name.

**Output**
- Output location: active sheet (next to existing data), a new worksheet (created once, reused), or a cell you pick.
- Output as an Excel table (with headers) or as a plain range with a defined name (headers off). Names are sanitised and made unique.
- Never overwrites non-empty cells: if the target is occupied the operation stops before writing anything.

**Chart**
- XY scatter with straight lines, markers off by default; Y axis fixed to 0–1 and formatted as percentages; axis titles; optional major gridlines; title; legend shown automatically for 2+ series.
- Start X configurable: automatic (0, or a round number below a negative minimum), zero, chart minimum, or custom.
- Final horizontal segment optionally extended to the chart maximum or a custom X.
- All distributions share the same start/end X so they compare on identical axes.

**Multiple distributions**
- Add any number of distributions; each gets its own coordinate table and chart series.
- Rename, hide/show on chart, remove (deletes only the generated table and series), select source/output ranges.
- **Regenerate** re-reads and re-validates every source range, rewrites the tables and refreshes the chart.
- State is stored in the workbook (Excel settings), so it survives closing the pane or reopening the file.

**Stochastic dominance analysis**
- Expected value, variance, standard deviation, min, median, max.
- First-order (FSD) and second-order (SSD) dominance with exact evaluation for discrete distributions; reports A dominates B, B dominates A, neither (with crossing points), or identical.
- “Higher is better” (profits) or “lower is better” (costs) preference.
- Plain-language interpretation; never classifies dominance from mean or variance.
- Optional export of the full comparison to a new worksheet.

---

## Requirements and compatibility

| Item | Requirement |
|---|---|
| Excel API | **ExcelApi 1.7** minimum (declared in both manifests). 1.8, 1.9 and 1.10 features are detected at run time with fallbacks. |
| Excel for Windows | Microsoft 365 (unified manifest: Version 2501+), Office 2021 / 2024 (XML manifest). Requires the WebView2 runtime. |
| Excel for Mac | Microsoft 365 / Excel 2021+; unified manifest needs **16.103 or later**. |
| Excel on the web | Supported (unified or XML manifest). |
| Not supported | Office 2019 and older on Windows (legacy EdgeHTML / IE webviews cannot run the React 19 / Fluent UI 9 bundle). |
| Node.js | 20 or later for building (developed with Node 26, npm 11). |

### Two manifests

| File | Format | Use it for |
|---|---|---|
| `manifest.json` | **Unified manifest for Microsoft 365** (schema 1.24, GA for Excel since July 2026) | Primary manifest. Microsoft 365 on Windows/Mac/web, admin-center deployment. |
| `manifest.xml` | Add-in only (XML) manifest | Perpetual Office 2021/2024, Mac builds older than 16.103, quick file-based sideloading without a Microsoft 365 sign-in. |

Both point to the same web files and declare the same ribbon button (**Home › Probability › Step Chart**).
They have different add-in IDs so they can be sideloaded side by side without cache conflicts.

### Office.js APIs used and their requirement sets

Checked against `@types/office-js` 1.0.612:

| API | Set | Notes |
|---|---|---|
| `workbook.getSelectedRange`, `worksheet.getRange`, `charts.add`, chart title / legend / axis min-max / axis titles / gridlines, `chart.setPosition`, `tables.add` | 1.1 | |
| `getUsedRange(OrNullObject)`, `getIntersectionOrNullObject`, `getItemOrNullObject`, `workbook.settings`, `names.add` | 1.2 – 1.4 | |
| `getRangeByIndexes`, `getSurroundingRegion`, `series.add/delete`, `setXAxisValues`, `setValues`, `series.chartType`, `markerStyle`, `smooth`, `format.line.weight` | **1.7** | manifest minimum |
| `ChartAxis.numberFormat`, `ChartAxis.position` | 1.8 | fallback: Y cells are formatted as % (axis follows its source format) |
| `workbook.getSelectedRanges` (multi-area detection) | 1.9 | fallback: Excel's own error for multi-area selections |
| `Range.left/top/width/height` (keep the chart off the data) | 1.10 | fallback: chart keeps its position |

Each chart formatting step is synced separately inside `try/catch`, so a host that rejects one property still produces a working chart and the pane lists what was skipped.

---

## Quick start (local development)

```bash
npm install
```

Trust the localhost development certificate once. This installs a certificate in your system trust store and asks for your password:

```bash
npm run certs
```

Build, start the HTTPS dev server on https://localhost:3000 and sideload into desktop Excel using the unified manifest:

```bash
npm start
```

Stop the server and unregister the add-in when you're done:

```bash
npm stop
```

Use the XML manifest instead (no Microsoft 365 sign-in, works with perpetual Office):

```bash
npm run start:xml
```

Try the UI in any browser without Excel (the pane runs against an in-memory workbook; see [Testing](#testing)):

```bash
npm run preview
```

Other scripts:

| Script | What it does |
|---|---|
| `npm run dev-server` | HTTPS dev server only (sideload manually). |
| `npm run build` | Production build into `dist/` (set `ADDIN_PUBLIC_URL`, see packaging). |
| `npm run watch` | Development build in watch mode. |
| `npm test` | Unit and integration tests (Vitest). |
| `npm run typecheck` | TypeScript type check. |
| `npm run validate` / `validate:xml` | Validate `manifest.json` against the unified schema / `manifest.xml` with Microsoft's validation service. |
| `npm run check` | Type check + tests + manifest validation. |
| `npm run package` | Build the Microsoft 365 app package zip from `dist/`. |

---

## Sideloading

Close Excel before sideloading. Keep `npm run dev-server` (or `npm start`) running while you use the add-in.

### Excel for Mac

**Automatic:** `npm start` (unified manifest; Excel 16.103+) or `npm run start:xml`.

**Manual, XML manifest (any Mac build with ExcelApi 1.7):**

1. Start the server: `npm run dev-server`.
2. Copy the manifest into Excel's sideload folder (create it if missing):

   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef && cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```

3. Open Excel, open a workbook, then **Home › Add-ins** (or **Insert › Add-ins › My Add-ins**) and choose **Probability Step Chart** under Developer Add-ins.

Remove it by deleting the copied file from the `wef` folder.

**Manual, unified manifest:** run `npm run package` after a build, then upload `dist/probability-step-chart.zip` via Teams (**Apps › Manage your apps › Upload an app › Upload a custom app**) with the same Microsoft 365 account you use in Excel, and restart Excel.

### Excel for Windows

**Automatic:** `npm start` (unified manifest, Microsoft 365 Version 2501+) or `npm run start:xml` (Office 2021/2024 or Microsoft 365).

**Manual, XML manifest via a shared-folder catalog:**

1. Share a folder (e.g. `C:\AddinManifests`) and copy `manifest.xml` into it.
2. Excel › **File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs**: add the share's network path (`\\machine\AddinManifests`), tick **Show in Menu**, and restart Excel.
3. **Home › Add-ins › More Add-ins › Shared Folder** › Probability Step Chart.

**Manual, unified manifest:** upload the app package zip through Teams as described for Mac.

### Excel on the web

- Unified manifest: sideloading on Windows also sideloads to Excel on the web for the same account; or run `npm run start:web` (add `-- --document <workbook URL>` to open a specific workbook).
- XML manifest: open a workbook at office.com, **Home › Add-ins › More Add-ins › My Add-ins › Upload My Add-in**, and choose `manifest.xml`.

---

## Using the add-in

Open `examples/ProbabilityStepChart-Examples.xlsx` for ready-made data (Make vs Buy, edge cases, dominance pairs and a reference chart).

1. **Build tab › Input:** select the two columns (outcome, probability), with or without a header row, and choose **Use selection**. The pane shows the address, header detection (override if needed), how probabilities were read, the sorted/merged distribution with cumulative probabilities, Σp, EV and SD, and any errors or warnings with cell references.
2. **Distribution name:** used as the chart series name (defaults to the text above the range).
3. **Step coordinates:** choose the start X and whether to extend the final segment; **Preview** shows the exact coordinates that will be written.
4. **Output:** location, table name, headers on/off.
5. **Chart:** title, axis labels, percentage Y axis, gridlines, markers. Untick **Create / update chart** to write only the table.
6. **Generate CDF coordinates.** Generating again from the same source range updates that distribution in place.
7. Select another range and generate again to **add a distribution to the same chart**.
8. **Distributions tab:** rename, show/hide on the chart, remove, jump to the source or output range, and **Regenerate all from source data** after editing inputs.
9. **Analysis tab:** pick distributions A and B and the preference (higher or lower is better). Read the statistics and the FSD/SSD verdicts, expand the CDF comparison table, or **Write to new sheet**.

---

## How the calculations work

All of this lives in `src/core/` and is independent of Office.js.

### CDF step coordinates (`cdf.ts`)

1. Sort outcomes ascending (duplicates already merged, zero-probability outcomes skipped).
2. Start at `(x_start, 0)`.
3. For each outcome `x` with probability `p`: add `(x, F_before)`, set `F ← F + p`, add `(x, F_after)`.
4. Snap the final cumulative value to exactly 1 when it is within tolerance (rejecting it otherwise); intermediate sums are rounded to 12 decimals so `0.1 + 0.2` is stored as `0.3`.
5. Optionally add `(x_end, 1)`.
6. Drop consecutive duplicate points (e.g. when `x_start` equals the first outcome).

**Automatic start:** 0 when every outcome is ≥ 0; otherwise a round number strictly below the minimum (e.g. −30 for outcomes from −20 to 50). When several distributions share a chart, the start and end are computed over all of them.

### Statistics (`stats.ts`)

`EV = Σ x·p`, `Var = Σ p·(x − EV)²` (population variance of the discrete distribution), `SD = √Var`. These are computed with compensated summation and normalised by Σp.

### Stochastic dominance (`dominance.ts`)

Let `z₁ < … < zₙ` be the union of the outcomes of A and B.

- **FSD.** Both CDFs are right-continuous step functions that change only at the `zₖ`, so `F_A − F_B` is constant on each `[zₖ, zₖ₊₁)`. Comparing `F_A(zₖ)` with `F_B(zₖ)` at every `k` is therefore an exact test for all real x, including the jumps. A FSD B iff `F_A ≤ F_B` everywhere with strict inequality somewhere: every decision maker with increasing utility weakly prefers A.
- **SSD.** `G(x) = ∫₋∞ˣ F(t) dt` is piecewise linear with breakpoints at the same `zₖ` (and `G_A − G_B` is constant beyond `zₙ`), so evaluating `G(zₖ) = Σ_{j<k} F(z_j)(z_{j+1} − z_j)` exactly at the breakpoints decides SSD over the full support. A SSD B iff `G_A ≤ G_B` everywhere with strict inequality somewhere: every risk-averse decision maker (increasing, concave utility) weakly prefers A.
- **Lower is better** (costs) is handled by applying the same tests to −X. Equivalently, this compares `F_A ≥ F_B` and `∫ₓ^∞ (1 − F)`.
- FSD implies SSD. When the CDFs cross, the pane lists the crossing points. Mean and variance are reported but **never** used to decide dominance.

Example (Make vs Buy, higher is better): the CDFs cross at x = 30, so neither has FSD. `G_Make − G_Buy` takes the values 0, 1.6, 1.6, 1.0, 1.0, −1.4, so neither has SSD either, even though Make has the higher EV. Treated as **costs**, Buy SSD-dominates Make.

---

## Architecture

```
src/
  core/                 pure TypeScript, no Office.js (unit tested)
    parse.ts            cell values → validated, sorted, merged distribution
    cdf.ts              step coordinates, shared X range
    stats.ts            EV, variance, SD, median
    dominance.ts        exact FSD / SSD + plain-language explanation
    a1.ts, numeric.ts   A1 references, Excel names, float hygiene
  excel/                Office.js layer
    rangeIO.ts          read selection / source ranges / output cell
    output.ts           write, rewrite and clear coordinate tables (conflict-safe)
    chart.ts            XY scatter chart creation, series sync, formatting with fallbacks
    service.ts          operations: add/update, regenerate, rename, remove
    storage.ts          workbook-persisted state
    analysisSheet.ts    export comparison to a worksheet
    support.ts          requirement-set checks, error messages
  taskpane/             React 19 + Fluent UI 9 task pane (Build / Distributions / Analysis)
  commands/             function file required by the XML manifest
  preview/              browser preview harness (dev only)
tests/                  Vitest: core unit tests + Excel-layer integration tests
  fakes/fakeExcel.ts    in-memory Excel API fake that enforces load()/sync()
manifest.json           unified manifest (primary)
manifest.xml            add-in only manifest (compatibility)
scripts/                icons, example workbook, app-package zip
examples/               ProbabilityStepChart-Examples.xlsx
docs/INTEGRATION_TESTING.md
```

Every operation validates everything it can (parsing, shared X range, step generation) **before** touching the workbook, checks the target cells are empty, and persists state even if a later step fails.

---

## Testing

```bash
npm test
```

75 automated tests:

- **Core unit tests** (`tests/cdf|parse|dominance|a1.test.ts`) cover each case in the specification: the Make and Buy examples, unsorted, duplicate and zero-probability outcomes, negative outcomes, percentages (formatted, text, percentage points), sums ≠ 1 and tolerance, single outcomes, floating-point precision, FSD, SSD with crossing CDFs, “higher mean and lower variance without SSD”, cost preference and identical distributions.
- **Excel-layer integration tests** (`tests/excelLayer.test.ts`) run the real `src/excel` code against `tests/fakes/fakeExcel.ts`. That in-memory fake **throws if a property is read before `load()` + `context.sync()`**, the most common Office.js bug. The tests cover selection reading and expansion, table and chart creation, the second series, shared X range, update-in-place, regenerate, conflict refusal, new-worksheet output, defined names, unique names, the ExcelApi 1.7 fallback, rename, hide, remove, persistence, chart recreation and the analysis export.

These do not replace testing in real Excel. Use the manual checklist in [`docs/INTEGRATION_TESTING.md`](docs/INTEGRATION_TESTING.md) for Excel on Mac, Windows and the web.

`npm run preview` serves the real task pane next to a live view of the in-memory workbook at http://localhost:8123, handy for UI work without Excel.

---

## Debugging

- **Mac:** enable the Web Inspector once:

  ```bash
  defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
  ```

  Then right-click in the task pane › **Inspect Element**.
- **Windows (WebView2):** right-click in the task pane › **Inspect**, or press F12 to open Edge DevTools. In VS Code you can attach with the *Microsoft Edge Tools* extension.
- **Web:** use the browser's DevTools; the pane runs in an iframe.
- Source maps are emitted for development builds (`npm run build:dev`, `npm run dev-server`).
- Errors from Excel are shown in the pane as “Excel reported: …”, with the failing API location when Office provides one.

---

## Packaging and distribution

1. **Host the files.** Build with the public HTTPS URL where `dist/` will be served (any static host: Azure Static Web Apps, S3 + CloudFront, GitHub Pages, IIS…):

   ```bash
   ADDIN_PUBLIC_URL=https://addins.example.com/stepchart/ npm run build
   ```

   Every `https://localhost:3000/` in the manifests and the about page is rewritten to that URL. Upload the whole `dist/` folder.

2. **Unified manifest app package:**

   ```bash
   npm run package
   ```

   This creates `dist/probability-step-chart.zip` (manifest.json + color/outline icons). The script warns if the manifest still points at localhost.

3. **Distribute:**
   - **Organisation-wide:** Microsoft 365 admin center › **Settings › Integrated apps › Upload custom apps** and upload the zip (unified) or `dist/manifest.xml`. Assign to users or groups. Deployment can take up to 24 hours.
   - **Individuals / testing:** Teams custom-app upload (zip), or the XML manifest via a shared-folder catalog (Windows), the `wef` folder (Mac), or Upload My Add-in (web).
   - **AppSource:** submit through Partner Center. Before submitting, replace the developer name and the website/privacy/terms URLs in both manifests with your own public pages (`src/static/about.html` provides a privacy and terms page you can host), bump `version`, and run `npm run validate` and `npm run validate:xml -- -p`.
4. **Updating:** web-file changes take effect on reload. Manifest changes (ribbon, permissions, version) need re-deployment.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Pane is blank or says “can’t reach localhost” | Is `npm run dev-server` running? Has the certificate been trusted (`npm run certs`)? Open https://localhost:3000/taskpane.html in a browser to check. |
| Ribbon button missing after sideloading | **Home › Add-ins** and pick the add-in once; restart Excel. |
| Old version keeps loading | Clear the Office cache. Mac: quit Excel and delete the contents of `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` and `.../Data/Library/Application Support/Microsoft/Office/16.0/Wef/`. Windows: delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. |
| “Unsupported Excel version” banner | Excel lacks ExcelApi 1.7; update Office. |
| “Output … needs D1:E9, but … are not empty” | Choose another output location or clear those cells. The add-in never overwrites data. |
| “… is no longer valid” on Regenerate | A source range now fails validation; the message names the cell. Fix the data and regenerate. |
| A chart formatting note appears | That Excel build rejected one formatting property; the chart still works. The note names the feature. |

---

## Known limitations

- When you switch the end-X option from “Chart maximum” or a custom value back to “Do not extend”, an existing chart keeps its previous X-axis maximum. Office.js cannot reset an axis to automatic, so set it in Excel or delete the chart and regenerate.
- Excel 2019 and earlier on Windows are not supported (see compatibility).
- Text numbers with thousands separators or locale-specific decimal commas (e.g. `"1,5"`) are not parsed. Enter numbers as numeric cells.
- Removing a distribution deletes only the add-in's generated table and series. If you have moved the generated table yourself, delete it manually.
