/** Export a stochastic-dominance comparison to a new worksheet. */
import { sanitizeExcelName, uniqueName } from "../core/a1";
import { DominanceExplanation, DominanceResult } from "../core/dominance";
import { cleanFloat } from "../core/numeric";
import { DistributionStats } from "../core/stats";
import { asText } from "./output";

export async function writeAnalysisSheet(
  nameA: string,
  nameB: string,
  statsA: DistributionStats,
  statsB: DistributionStats,
  result: DominanceResult,
  explanation: DominanceExplanation,
): Promise<string> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();
    const taken = new Set(sheets.items.map((s) => s.name.toLowerCase()));
    const tableNames = tables.items.map((t) => t.name);
    let sheetName = "Dominance Analysis";
    for (let i = 2; taken.has(sheetName.toLowerCase()); i++) sheetName = `Dominance Analysis ${i}`;
    const sheet = sheets.add(sheetName);

    const higher = result.preference === "higher";
    const intLabel = higher ? "∫₋∞ˣ F(t) dt" : "∫ₓ^∞ (1−F(t)) dt";
    const A = asText(nameA);
    const B = asText(nameB);
    const header: (string | number)[][] = [
      ["Stochastic dominance analysis", "", ""],
      ["Preference", higher ? "Higher outcomes are better" : "Lower outcomes are better (e.g. costs)", ""],
      ["", "", ""],
      ["Statistic", A, B],
      ["Expected value", statsA.expectedValue, statsB.expectedValue],
      ["Variance", statsA.variance, statsB.variance],
      ["Standard deviation", statsA.standardDeviation, statsB.standardDeviation],
      ["Minimum", statsA.min, statsB.min],
      ["Median", statsA.median, statsB.median],
      ["Maximum", statsA.max, statsB.max],
      ["", "", ""],
      ["First-order (FSD)", explanation.fsdHeadline, ""],
      ["", explanation.fsdDetail, ""],
      ["Second-order (SSD)", explanation.ssdHeadline, ""],
      ["", explanation.ssdDetail, ""],
      ...explanation.notes.map((n, i) => [i === 0 ? "Notes" : "", n, ""]),
      ["", "", ""],
    ];
    sheet.getRangeByIndexes(0, 0, header.length, 3).values = header;
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 14;
    sheet.getRange("A4:C4").format.font.bold = true;
    sheet.getRangeByIndexes(11, 0, 4, 1).format.font.bold = true;
    sheet.getRangeByIndexes(11, 1, 1, 1).format.font.bold = true;
    sheet.getRangeByIndexes(13, 1, 1, 1).format.font.bold = true;

    const start = header.length;
    const rows = result.rows.map((r) => [
      r.x,
      r.FA,
      r.FB,
      cleanFloat(r.FA - r.FB),
      r.IA,
      r.IB,
      cleanFloat(r.IA - r.IB, 10),
    ]);
    const columns = ["x", `F ${A}`, `F ${B}`, "F difference (A − B)", `${intLabel} ${A}`, `${intLabel} ${B}`, "Integral difference (A − B)"];
    sheet.getRangeByIndexes(start, 0, rows.length + 1, columns.length).values = [columns, ...rows];
    const table = sheet.tables.add(sheet.getRangeByIndexes(start, 0, rows.length + 1, columns.length), true);
    table.name = uniqueName(sanitizeExcelName("Dominance_Comparison"), tableNames);
    sheet.getRangeByIndexes(start + 1, 1, rows.length, 3).numberFormat = rows.map(() => ["0.0%", "0.0%", "0.0%"]);
    sheet.getRangeByIndexes(0, 0, 1, 1).format.columnWidth = 120;
    sheet.getRangeByIndexes(0, 1, 1, columns.length - 1).format.columnWidth = 110;
    sheet.activate();
    await context.sync();
    return sheetName;
  });
}
