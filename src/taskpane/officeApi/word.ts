/** Word insertion helpers — full-fidelity host: Word.js supports rich text
 * and native table insertion, but no chart-object API, so charts fall back
 * to an inserted image (see insertChartImage). */

import type { XlsxTable } from "../../chatClient";

export async function insertProse(text: string): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(text, Word.InsertLocation.after);
    await context.sync();
  });
}

export async function insertTable(table: XlsxTable): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const rowCount = table.rows.length + 1; // + header row
    const colCount = table.columns.length;
    const wordTable = range.insertTable(
      rowCount,
      colCount,
      Word.InsertLocation.after,
      [table.columns.map(String)]
    );
    for (let r = 0; r < table.rows.length; r++) {
      for (let c = 0; c < colCount; c++) {
        const cell = wordTable.getCell(r + 1, c);
        cell.body.insertText(String(table.rows[r][c] ?? ""), Word.InsertLocation.replace);
      }
    }
    wordTable.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent1;
    await context.sync();
  });
}

/** Word.js has no native chart-object creation, so a requested chart is
 * rendered server-side (`/chat/charts/render`) and dropped in as a static
 * inline picture — not live/editable, but placed at the cursor like any
 * other figure. */
export async function insertChartImage(pngBase64: string): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertInlinePictureFromBase64(pngBase64, Word.InsertLocation.after);
    await context.sync();
  });
}
