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

export async function insertReport(answer: any): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const anchor = context.workbook.getSelectedRange();
    anchor.load("rowIndex, rowCount");
    await context.sync();
    
    sheet.activate();
    // Start inserting 1 row below the bottom of the current selection
    let currentRow = anchor.rowIndex + anchor.rowCount;
    const tableAddresses = new Map<number, string>();

    const operations = answer.plan?.operations || [];
    for (const op of operations) {
      if (op.type === "title" || op.type === "header") {
        const range = sheet.getRangeByIndexes(currentRow, 1, 1, 8); // B:I
        range.values = [[op.text, "", "", "", "", "", "", ""]];
        range.merge();
        range.format.font.bold = true;
        range.format.font.size = 20;
        range.format.font.color = "#111827";
        currentRow += 2;
      } else if (op.type === "subtitle") {
        const range = sheet.getRangeByIndexes(currentRow, 1, 1, 8);
        range.values = [[op.text, "", "", "", "", "", "", ""]];
        range.merge();
        range.format.font.size = 14;
        range.format.font.color = "#4b5563";
        currentRow += 2;
      } else if (op.type === "heading") {
        const range = sheet.getRangeByIndexes(currentRow, 1, 1, 8);
        range.values = [[op.text, "", "", "", "", "", "", ""]];
        range.merge();
        range.format.font.bold = true;
        range.format.font.size = 14 + (6 - (op.level || 3));
        currentRow += 2;
      } else if (op.type === "paragraph" || op.type === "bullet" || op.type === "numbered") {
        const range = sheet.getRangeByIndexes(currentRow, 1, 1, 8);
        range.values = [[op.text, "", "", "", "", "", "", ""]];
        range.merge();
        range.format.wrapText = true;
        range.format.font.size = 11;
        range.format.autofitRows();
        currentRow += 1;
      } else if (op.type === "table" && op.table_index !== null) {
        currentRow += 1;
        const table = answer.tables[op.table_index];
        if (table) {
          const rowCount = table.rows.length + 1;
          const colCount = table.columns.length;
          const range = sheet.getRangeByIndexes(currentRow, 1, rowCount, colCount);
          
          range.values = [
            table.columns.map((c: any) => String(c).replace(/^\*\*(.*)\*\*$/, "$1").trim()),
            ...table.rows.map((r: any) =>
              r.map((v: any) => {
                if (v === null || v === undefined) return "";
                if (typeof v === "number" || typeof v === "boolean") return v;
                return String(v).replace(/^\*\*(.*)\*\*$/, "$1").trim();
              })
            ),
          ];
          
          const excelTable = sheet.tables.add(range, true);
          excelTable.name = `ReportTable_${Date.now()}_${op.table_index}`;
          
          let hasTotalRow = false;
          // Apply format from layout panel
          if (answer.layout?.panels) {
            const panel = answer.layout.panels.find((p: any) => p.table_index === op.table_index);
            if (panel && panel.columns) {
              for (const colFmt of panel.columns) {
                if (colFmt.index < colCount) {
                  const colRange = excelTable.columns.getItemAt(colFmt.index).getDataBodyRange();
                  if (colFmt.number_format) {
                    colRange.numberFormat = [[colFmt.number_format]];
                  }
                  if (colFmt.align) {
                    colRange.format.horizontalAlignment = colFmt.align === "right" ? "Right" : colFmt.align === "center" ? "Center" : "Left";
                  }
                  if (colFmt.bar) {
                    const condFmt = colRange.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                    condFmt.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
                  }
                }
              }
              if (panel.has_total_row && table.rows.length > 0) {
                hasTotalRow = true;
                const lastRowRange = excelTable.getDataBodyRange().getLastRow();
                lastRowRange.format.font.bold = true;
                lastRowRange.format.borders.getItem("EdgeTop").style = "Double";
              }
            }
          }
          
          range.format.autofitColumns();
          const chartRange = hasTotalRow ? sheet.getRangeByIndexes(currentRow, 1, rowCount - 1, colCount) : range;
          chartRange.load("address");
          await context.sync();
          tableAddresses.set(op.table_index, chartRange.address);
          
          currentRow += rowCount + 1;
        }
      } else if (op.type === "chart" && op.table_index !== null) {
        currentRow += 1;
        const address = tableAddresses.get(op.table_index);
        if (address) {
          const chartTypeMap: Record<string, any> = {
            bar: Excel.ChartType.columnClustered,
            line: Excel.ChartType.line,
            area: Excel.ChartType.area,
            pie: Excel.ChartType.pie,
            donut: Excel.ChartType.doughnut,
          };
          const kind = chartTypeMap[op.chart_type] ?? Excel.ChartType.columnClustered;
          
          const dataRange = sheet.getRange(address);
          const chart = sheet.charts.add(kind, dataRange, Excel.ChartSeriesBy.columns);
          
          // Position chart below the current row
          const chartTopLeft = sheet.getRangeByIndexes(currentRow, 1, 1, 1);
          const chartBottomRight = sheet.getRangeByIndexes(currentRow + 15, 8, 1, 1);
          chart.setPosition(chartTopLeft, chartBottomRight);
          
          currentRow += 16;
        }
      }
    }
    
    await context.sync();
  });
}
