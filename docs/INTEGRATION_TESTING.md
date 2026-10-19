# Integration testing checklist

Run this manually in real Excel before each release. The automated suite (`npm test`)
covers the math and runs the Excel layer against an in-memory fake; this checklist
covers what only real Excel can show: rendering, host-specific API behaviour, sideloading
and performance.

Use `examples/ProbabilityStepChart-Examples.xlsx`. Record results in the matrix at the end.

## Environments

| ID | Platform | Manifest | Notes |
|---|---|---|---|
| M1 | Excel for Mac, Microsoft 365, ≥ 16.103 | `manifest.json` (`npm start`) | |
| M2 | Excel for Mac | `manifest.xml` (wef folder) | |
| W1 | Excel for Windows, Microsoft 365 ≥ 2501 | `manifest.json` (`npm start`) | WebView2 |
| W2 | Excel 2021 / 2024 perpetual, Windows | `manifest.xml` (shared folder) | |
| X1 | Excel on the web (Edge or Chrome) | either | |
| X2 | Excel on the web (Safari) | either | optional |

## A. Installation

- [ ] A1. The add-in loads and **Home › Probability › Step Chart** appears on the ribbon.
- [ ] A2. Opening the pane shows the Build / Distributions / Analysis tabs with no error banner.
- [ ] A3. In a dark Office theme (Windows/Mac), the pane uses the dark theme.
- [ ] A4. Closing and reopening the pane in the same workbook restores the distribution list and chart settings.
- [ ] A5. Saving, closing and reopening the workbook restores the distribution list (state stored in the workbook).

## B. Input and validation (sheet “Make vs Buy” and “Edge Cases”)

- [ ] B1. Select `A2:B5` › **Use selection** › the address is shown, header found, name “Make Component”, preview 11/32/53 with cumulative 40/70/100 %, Σp = 1, EV 29.9.
- [ ] B2. Click only `A3` › Use selection › expands to `A2:B5` with the info message.
- [ ] B3. Select whole columns `A:B` on a sheet with one table › trimmed to the data.
- [ ] B4. Ctrl/Cmd-select two separate areas › clear “multiple areas” message (ExcelApi 1.9+).
- [ ] B5. Header override “No header row” on `A2:B5` › error at `A2` (“Outcome … is not a number”).
- [ ] B6. Edge Cases › Unsorted: “sorted” info. Duplicate: warning and merged 0.4. Zero probability: warning.
- [ ] B7. Percent-formatted cells, `"40%"` text and 40/30/30 points all give 0.4/0.3/0.3.
- [ ] B8. Invalid sum 0.95: error, **Generate** disabled. Text outcome: error names the cell. Probability 1.2 / −0.2: two errors.
- [ ] B9. Tolerance ±1e−2 lets `0.333/0.333/0.333` pass; ±1e−6 rejects it.
- [ ] B10. The source range is never modified by any action in this checklist.

## C. Coordinates and table output

- [ ] C1. Generate Make (active sheet) › table placed one blank column right of existing data, label row “Make Component”, headers “X (Outcome)” / “Y (Cumulative Probability)”, rows exactly as on the “Reference Output” sheet.
- [ ] C2. The table is a real Excel table named `CDF_Make_Component` (Table Design › Table Name).
- [ ] C3. Untick **Include headers** and generate another range › plain data with a defined name (Formulas › Name Manager).
- [ ] C4. Output **New worksheet** › sheet “CDF Output” created; a second distribution goes beside the first.
- [ ] C5. Output **Selected cell** on an occupied cell › refusal message, nothing written. On an empty cell › written there.
- [ ] C6. Generate the same source again with a new name › updated in place (no second table).
- [ ] C7. Edge Cases › Negative outcomes › automatic start X is −30, and the first row is (−30, 0).
- [ ] C8. Start X **Custom** = 20 with Make › clear error, nothing written.
- [ ] C9. X values keep the source number format (e.g. format the Make outcomes as currency first).

## D. Chart

- [ ] D1. Chart type is XY scatter with straight lines (Chart Design › Change Chart Type shows “Scatter with Straight Lines”).
- [ ] D2. Steps are drawn as horizontal segments with vertical jumps; no markers.
- [ ] D3. Y axis 0 %–100 %, major unit 10 %, title “Cumulative Probability”. X axis title “Outcome”.
- [ ] D4. Gridlines toggle works; markers toggle shows markers.
- [ ] D5. Chart title text and blank title (hidden) both work.
- [ ] D6. Add Buy (`D2:E5`) › second series on the same chart, legend at the bottom, Buy extended to X = 53.
- [ ] D7. Negative data › Y axis sits at the left edge, not at X = 0 (ExcelApi 1.8+).
- [ ] D8. Move the chart over the coordinate tables, then add another distribution › chart moves off the data (ExcelApi 1.10+).
- [ ] D9. Recolour a series manually, then Regenerate › the colour is kept.
- [ ] D10. Delete the chart manually, then Regenerate › chart recreated.
- [ ] D11. No formatting notes appear on current builds. On older builds, any note names a skipped feature and the chart still works.

## E. Multiple distributions

- [ ] E1. Rename Buy to “Outsource” in the Distributions tab › series name and label cell update.
- [ ] E2. **Show on chart** off › series removed, table kept. On › series back.
- [ ] E3. Change `B3` from 0.4 to 0.5 and `B4` from 0.3 to 0.2 › **Regenerate** › Make table and chart update.
- [ ] E4. Make a source invalid, then Regenerate › error names the distribution and the cell; tables unchanged.
- [ ] E5. Source/Output links select the right ranges, including on other sheets.
- [ ] E6. Rename the source worksheet › Regenerate still works (sheets are tracked by ID).
- [ ] E7. Remove a distribution › its table is cleared, its series removed, the source untouched. Removing the last one deletes the chart.

## F. Stochastic dominance (sheet “Dominance” and Make vs Buy)

Add both distributions of each pair, then open **Analysis**.

- [ ] F1. Make vs Buy, higher is better › FSD: neither (crossing near 30); SSD: neither; EV 29.9 / 28.5; Var 304.29 / 155.25.
- [ ] F2. Make vs Buy, **lower is better** › FSD: neither; SSD: Buy dominates.
- [ ] F3. Pair “A FSD B” › FSD: A dominates; SSD: A dominates (implied).
- [ ] F4. Pair “sure 10 vs 0/20” › FSD: neither; SSD: A dominates.
- [ ] F5. Pair “higher mean, lower variance” › neither; the note explains that the higher EV does not imply dominance.
- [ ] F6. Pair “identical” › identical.
- [ ] F7. **Write to new sheet** › “Dominance Analysis” sheet with statistics, verdicts and a comparison table; running it again creates “Dominance Analysis 2”.

## G. Robustness and performance

- [ ] G1. A distribution with 1,000 outcomes generates in a few seconds and the chart renders.
- [ ] G2. Undo (Ctrl/Cmd+Z) after Generate behaves as Excel allows (add-in actions may clear the undo stack; note the behaviour).
- [ ] G3. The pane works when the active sheet is protected elsewhere: writing to a protected sheet shows Excel's error message, not a crash.
- [ ] G4. Resizing the task pane to its narrowest width causes no horizontal scrolling.

## Results matrix

| Section | M1 | M2 | W1 | W2 | X1 | X2 |
|---|---|---|---|---|---|---|
| A Installation | | | | | | |
| B Input | | | | | | |
| C Output | | | | | | |
| D Chart | | | | | | |
| E Multiple | | | | | | |
| F Dominance | | | | | | |
| G Robustness | | | | | | |

Tester / date / Excel build numbers:
