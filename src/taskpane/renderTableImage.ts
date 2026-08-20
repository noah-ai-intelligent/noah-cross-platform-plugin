/** Client-rendered table → PNG (base64, no data: prefix) — PowerPoint's JS
 * API can't insert a native table, so a plain table (no chart requested)
 * falls back to a simple canvas rendering, same idea as the server-rendered
 * chart image used for charts. Deliberately basic: this is a fallback, not
 * the primary path (Word/Excel/Docs/Sheets/Slides all get real tables). */

import type { XlsxTable } from "../chatClient";

const ROW_H = 28;
const PAD = 10;
const FONT = "13px -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif";

export function renderTableImageBase64(table: XlsxTable): string {
  const rows = table.rows.length;

  const measure = document.createElement("canvas").getContext("2d")!;
  measure.font = FONT;
  const colWidths = table.columns.map((c, i) => {
    const headerW = measure.measureText(String(c)).width;
    const cellW = Math.max(
      0,
      ...table.rows.map((r) => measure.measureText(String(r[i] ?? "")).width)
    );
    return Math.max(headerW, cellW) + PAD * 2;
  });
  const width = Math.max(200, colWidths.reduce((a, b) => a + b, 0));
  const height = ROW_H * (rows + 1);

  const canvas = document.createElement("canvas");
  canvas.width = width * 2; // 2x for crisper text when inserted at slide size
  canvas.height = height * 2;
  const ctx = canvas.getContext("2d")!;
  ctx.scale(2, 2);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.font = FONT;
  ctx.textBaseline = "middle";

  let x = 0;
  ctx.fillStyle = "#0d5c63";
  ctx.fillRect(0, 0, width, ROW_H);
  ctx.fillStyle = "#ffffff";
  table.columns.forEach((c, i) => {
    ctx.fillText(String(c), x + PAD, ROW_H / 2, colWidths[i] - PAD * 2);
    x += colWidths[i];
  });

  table.rows.forEach((r, ri) => {
    const y = ROW_H * (ri + 1);
    ctx.fillStyle = ri % 2 === 0 ? "#f1f5f9" : "#ffffff";
    ctx.fillRect(0, y, width, ROW_H);
    ctx.fillStyle = "#1e293b";
    let cx = 0;
    r.forEach((v, ci) => {
      ctx.fillText(String(v ?? ""), cx + PAD, y + ROW_H / 2, colWidths[ci] - PAD * 2);
      cx += colWidths[ci];
    });
  });

  return canvas.toDataURL("image/png").split(",")[1] ?? "";
}
