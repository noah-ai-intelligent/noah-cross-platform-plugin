/** Reading what the user has selected, in whichever host we're in.
 *
 * Deliberately dumb: capture and post. Every decision about what to KEEP —
 * trimming blank rows, typing "1,234.50" into a number, detecting a header row
 * or a total line, capping the size — is made server-side in
 * `app/api/v1/addons/normalize.py`, once, for all six hosts. Six clients each
 * deciding what a "selection" means is six drift surfaces and six token-budget
 * bugs.
 *
 * So this file sends the raw rectangle, including the trailing blank columns
 * the user happened to drag over. The server knows what to do with them.
 */

import {
  SelectionAnchor,
  gridAnchor,
  hashText,
  slideAnchor,
  textAnchor,
} from "./anchor";
import { Host, currentHost } from "./capabilities";

export interface GridPayload {
  columns: string[];
  rows: unknown[][];
  number_formats?: string[];
  merged?: string[];
}

export interface TextPayload {
  blocks: { type: string; text: string; level?: number }[];
  outline?: string[];
  char_count: number;
}

export interface SlideCapture {
  slide_id: string;
  index: number;
  title: string;
  body: string[];
  notes: string;
}

export interface Selection {
  anchor: SelectionAnchor;
  grid?: GridPayload;
  text?: TextPayload;
  shapes?: { slides: SlideCapture[] };
}

/** How many cells we're willing to pull out of the host in one go. Well above
 * anything a person selects by hand; it exists so "select all" on a big sheet
 * doesn't lock the taskpane. The server trims further to policy. */
const MAX_CELLS = 200_000;

async function captureExcel(): Promise<Selection | null> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("address,values,numberFormat,cellCount,rowCount,columnCount");
    await ctx.sync();

    if (!range.rowCount || !range.columnCount) return null;
    if (range.cellCount > MAX_CELLS) {
      throw new Error(
        `That selection is ${range.cellCount.toLocaleString()} cells — select a smaller range.`
      );
    }

    const values = (range.values ?? []) as unknown[][];
    // A single cell is a legitimate selection but rarely a useful question on
    // its own; send it anyway and let the agent decide.
    const formats = (range.numberFormat ?? []) as string[][];
    return {
      anchor: gridAnchor(range.address, values),
      grid: {
        // No header guess here — the server detects it, and a wrong guess made
        // client-side would be baked in before it got the chance.
        columns: [],
        rows: values,
        number_formats: formats.length ? formats[0].map((f) => String(f ?? "")) : [],
      },
    };
  });
}

async function captureWord(): Promise<Selection | null> {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text");
    const paras = sel.paragraphs;
    paras.load("items/text,items/style,items/styleBuiltIn");
    await ctx.sync();

    const text = (sel.text ?? "").trim();
    if (!text) return null;

    const blocks = paras.items.map((p) => {
      const style = String(p.styleBuiltIn ?? p.style ?? "");
      const heading = /^Heading(\d)$/.exec(style);
      if (heading) {
        return { type: "heading", text: p.text ?? "", level: Number(heading[1]) };
      }
      if (/List/.test(style)) return { type: "bullet", text: p.text ?? "" };
      return { type: "paragraph", text: p.text ?? "" };
    });

    // Word gives no stable character offsets for an arbitrary selection, so the
    // anchor carries the hash and no offsets. That is enough for a citation
    // chip (which re-selects by search) and for change detection; it is NOT
    // enough for an anchored edit, which is why C4 adds a bookmark here.
    return {
      anchor: { kind: "text", content_sha: hashText(text) },
      text: { blocks, char_count: text.length },
    };
  });
}

async function capturePowerPoint(): Promise<Selection | null> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id,items/index");
    await ctx.sync();

    if (!slides.items.length) return null;

    const captured: SlideCapture[] = [];
    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/name,items/textFrame/textRange/text");
      await ctx.sync();

      const texts = shapes.items
        .map((s) => {
          try {
            return s.textFrame?.textRange?.text ?? "";
          } catch {
            // A shape with no text frame (a picture, a line) throws rather
            // than returning empty.
            return "";
          }
        })
        .filter((t) => t.trim());

      captured.push({
        slide_id: String(slide.id ?? ""),
        index: Number(slide.index ?? 0),
        title: texts[0] ?? "",
        body: texts.slice(1),
        notes: "",
      });
    }

    const flat = captured
      .map((s) => [s.title, ...s.body].join("\n"))
      .join("\n");
    const first = captured[0];
    return {
      anchor: slideAnchor(first.index, first.slide_id, flat),
      shapes: { slides: captured },
    };
  });
}

/** The current selection, or null when there is nothing selected worth sending.
 *
 * Never throws for "nothing selected" — that is the normal case for someone
 * asking a question about org data, and the ask proceeds without context. It
 * DOES throw when the selection is unusably large, because silently sending
 * less than the user picked would produce an answer about data they didn't
 * choose.
 */
export async function captureSelection(
  host: Host = currentHost()
): Promise<Selection | null> {
  try {
    if (host === "Excel") return await captureExcel();
    if (host === "Word") return await captureWord();
    if (host === "PowerPoint") return await capturePowerPoint();
    return null;
  } catch (err) {
    if (err instanceof Error && /select a smaller range/.test(err.message)) throw err;
    // Any other host error means we couldn't read the selection. Asking
    // without it beats failing the question.
    console.warn("Noah: could not read the selection", err);
    return null;
  }
}

export { textAnchor };
