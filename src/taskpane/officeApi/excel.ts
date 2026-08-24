/** Excel insertion helpers — the richest host: native Tables AND native,
 * live/editable Charts both come straight from Excel.js, no image round-trip
 * needed for either. */

import type { XlsxTable } from "../../chatClient";

// CHART_TYPE_MAP removed from top level to avoid crash outside Excel host

/** Writes text directly into the selected cell/range in Excel. */
export async function insertProse(text: string): Promise<void> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.values = [[text]];
    await context.sync();
  });
}

/** Writes the table starting at the active cell and formats it as a native
 * Excel Table (sortable/filterable, matches the header styling). Returns the
 * written range so a chart request can build straight off it. */
export async function insertTable(
  table: XlsxTable,
  isUpdate: boolean = false
): Promise<{ address: string }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const anchor = context.workbook.getSelectedRange();
    anchor.load("address, rowIndex, columnIndex, rowCount");
    await context.sync();

    const rowCount = table.rows.length + 1;
    const colCount = table.columns.length;
    
    // If it's an update, overwrite exactly at the selection's start.
    // If it's a create, place it 1 cell below the bottom of the current selection.
    const startRow = isUpdate ? anchor.rowIndex : anchor.rowIndex + anchor.rowCount;
    
    const range = sheet.getRangeByIndexes(
      startRow,
      anchor.columnIndex,
      rowCount,
      colCount
    );

    // If it's an update, we want to clear the original selection too.
    // Otherwise, just clear the destination where the new table will go.
    const targetArea = isUpdate ? anchor.getBoundingRect(range) : range;

    sheet.tables.load("items");
    await context.sync();
    
    // ALWAYS check for intersecting tables on the target area and delete them
    for (const t of sheet.tables.items) {
      const tRange = t.getRange();
      const intersection = tRange.getIntersectionOrNullObject(targetArea);
      intersection.load("isNullObject");
      tRange.load("address");
      await context.sync();
      
      if (!intersection.isNullObject) {
        const address = tRange.address;
        t.delete();
        sheet.getRange(address).clear();
      }
    }
    
    targetArea.clear();
    await context.sync();

    range.values = [
      table.columns.map((c) => String(c).replace(/^\*\*(.*)\*\*$/, "$1").trim()),
      ...table.rows.map((r) =>
        r.map((v) => {
          if (v === null || v === undefined) return "";
          if (typeof v === "number" || typeof v === "boolean") return v;
          return String(v).replace(/^\*\*(.*)\*\*$/, "$1").trim();
        })
      ),
    ];
    const excelTable = sheet.tables.add(range, true /* hasHeaders */);
    excelTable.name = `NoahTable${Date.now()}`;
    range.format.autofitColumns();
    range.load("address");
    await context.sync();
    return { address: range.address };
  });
}

/** Builds a real, editable native chart directly from the range Excel just
 * wrote — no image round-trip, unlike Word/PowerPoint. */
export async function insertChart(
  chartType: string,
  rangeAddress: string
): Promise<void> {
  const chartTypeMap: Record<string, any> = {
    bar: Excel.ChartType.columnClustered,
    line: Excel.ChartType.line,
    area: Excel.ChartType.area,
    pie: Excel.ChartType.pie,
    donut: Excel.ChartType.doughnut,
  };
  const kind = chartTypeMap[chartType] ?? Excel.ChartType.columnClustered;
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const dataRange = sheet.getRange(rangeAddress);
    const chart = sheet.charts.add(kind, dataRange, Excel.ChartSeriesBy.columns);
    chart.title.text = "";
    await context.sync();
  });
}
