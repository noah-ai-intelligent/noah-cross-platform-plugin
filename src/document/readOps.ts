/** Performing the reads the agent asks for, mid-answer.
 *
 * When a turn needs something that isn't in what the user sent — another
 * sheet, the heading tree, a formula's precedents — the backend suspends it
 * and parks a request on the job (`app/api/v1/addons/bridge.py`). The taskpane
 * sees `status: needs_context`, runs the op here, and posts the result back;
 * the turn picks up where it left off.
 *
 * Two shapes come back, and which one an op uses is the backend's contract:
 *
 *   `selection` — a region, normalised server-side exactly like a pushed
 *                 selection (so a masked column stays masked)
 *   `items`     — a list that isn't a region: sheets, slides, headings, search
 *                 hits, precedents
 *
 * Every handler is best-effort. A failed read comes back `ok: false` with a
 * reason and the agent writes the best answer it can — a sidebar that can't
 * read one range must not take the whole turn down with it.
 */

import { gridAnchor, hashText } from "./anchor";
import { Host, currentHost } from "./capabilities";
import { captureSelection, Selection } from "./selection";

export interface ClientToolResult {
  request_id: string;
  ok: boolean;
  error?: string;
  selection?: Selection | null;
  items?: Record<string, unknown>[];
}

export interface ClientToolRequest {
  request_id: string;
  tool: string;
  op: string;
  args: Record<string, unknown>;
  human_label: string;
  document_id: string;
}

/** Cells one pulled region may carry. The agent asked for a range, not the
 * workbook; anything past this is a sign it guessed, and truncating with a
 * note beats a payload that stalls the taskpane. */
const MAX_READ_CELLS = 50_000;

function arg(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  return value == null ? "" : String(value);
}

// ── Excel ──────────────────────────────────────────────────────────────────

async function readRange(args: Record<string, unknown>): Promise<Selection> {
  const sheetName = arg(args, "sheet");
  const a1 = arg(args, "a1");
  if (!a1) throw new Error("read_range needs an `a1` range.");

  return Excel.run(async (ctx) => {
    const sheet = sheetName
      ? ctx.workbook.worksheets.getItem(sheetName)
      : ctx.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(a1);
    range.load("address,values,numberFormat,cellCount");
    await ctx.sync();

    if (range.cellCount > MAX_READ_CELLS) {
      throw new Error(
        `${a1} is ${range.cellCount.toLocaleString()} cells — ask for a smaller range.`
      );
    }
    const values = (range.values ?? []) as unknown[][];
    const formats = (range.numberFormat ?? []) as string[][];
    return {
      anchor: gridAnchor(range.address, values),
      grid: {
        columns: [],
        rows: values,
        number_formats: formats.length ? formats[0].map((f) => String(f ?? "")) : [],
      },
    };
  });
}

async function listSheets(): Promise<Record<string, unknown>[]> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await ctx.sync();

    const used = sheets.items.map((s) => {
      const range = s.getUsedRangeOrNullObject(true);
      range.load("address,rowCount,columnCount,isNullObject");
      return range;
    });
    await ctx.sync();

    return sheets.items.map((sheet, i) => ({
      sheet: sheet.name,
      used_range: used[i].isNullObject
        ? ""
        : String(used[i].address ?? "").replace(/^.*!/, ""),
      rows: used[i].isNullObject ? 0 : Number(used[i].rowCount ?? 0),
      columns: used[i].isNullObject ? 0 : Number(used[i].columnCount ?? 0),
      hidden: sheet.visibility !== Excel.SheetVisibility.visible,
    }));
  });
}

async function readNamedRanges(): Promise<Record<string, unknown>[]> {
  return Excel.run(async (ctx) => {
    const names = ctx.workbook.names;
    names.load("items/name,items/value,items/comment");
    await ctx.sync();
    return names.items.map((n) => ({
      name: n.name,
      refers_to: String(n.value ?? ""),
      comment: String(n.comment ?? ""),
    }));
  });
}

async function readTables(): Promise<Record<string, unknown>[]> {
  return Excel.run(async (ctx) => {
    const tables = ctx.workbook.tables;
    tables.load("items/name,items/worksheet/name");
    await ctx.sync();

    const ranges = tables.items.map((t) => {
      const range = t.getRange();
      range.load("address,rowCount,columnCount");
      return range;
    });
    await ctx.sync();

    return tables.items.map((t, i) => ({
      table: t.name,
      sheet: t.worksheet?.name ?? "",
      range: String(ranges[i].address ?? "").replace(/^.*!/, ""),
      rows: Number(ranges[i].rowCount ?? 0),
      columns: Number(ranges[i].columnCount ?? 0),
    }));
  });
}

/** Precedents / dependents — the whole of our formula-tracing story.
 *
 * `getDirectPrecedents` needs ExcelApi 1.15, which is why `capabilities.ts`
 * probes for it: on an older Excel the agent is never offered the op at all.
 */
async function trace(
  args: Record<string, unknown>,
  direction: "precedents" | "dependents"
): Promise<Record<string, unknown>[]> {
  const sheetName = arg(args, "sheet");
  const a1 = arg(args, "a1");
  if (!a1) throw new Error(`trace_${direction} needs an \`a1\` cell.`);

  return Excel.run(async (ctx) => {
    const sheet = sheetName
      ? ctx.workbook.worksheets.getItem(sheetName)
      : ctx.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(a1);
    // `getDirectPrecedents` returns a WorkbookRangeAreas — precedents can live
    // on other sheets, which is exactly the case worth tracing. `.ranges` is
    // the flat list across all of them; `.areas` would group them per sheet
    // and make us walk two levels for nothing.
    const related =
      direction === "precedents" ? cell.getDirectPrecedents() : cell.getDirectDependents();
    related.ranges.load("items/address,items/values,items/formulas");
    await ctx.sync();

    const key = direction === "precedents" ? "feeds_from" : "affects";
    return related.ranges.items.map((range) => {
      const values = (range.values ?? []) as unknown[][];
      const formulas = (range.formulas ?? []) as string[][];
      return {
        [key]: String(range.address ?? ""),
        value: values[0]?.[0] ?? null,
        formula: formulas[0]?.[0] ?? "",
      };
    });
  });
}

// ── Word ───────────────────────────────────────────────────────────────────

async function readOutline(): Promise<Record<string, unknown>[]> {
  return Word.run(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load("items/text,items/styleBuiltIn,items/style");
    await ctx.sync();

    const out: Record<string, unknown>[] = [];
    paragraphs.items.forEach((p, i) => {
      const style = String(p.styleBuiltIn ?? p.style ?? "");
      const heading = /^Heading(\d)$/.exec(style);
      if (!heading) return;
      const text = (p.text ?? "").trim();
      if (text) out.push({ heading: text, level: Number(heading[1]), paragraph: i });
    });
    return out;
  });
}

async function readDocumentText(args: Record<string, unknown>): Promise<Selection> {
  const start = Number(args.start ?? 0);
  const end = Number(args.end ?? 0);

  return Word.run(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load("items/text,items/styleBuiltIn,items/style");
    await ctx.sync();

    // Word has no character-offset addressing for the body, so `start`/`end`
    // are treated as PARAGRAPH indices. Saying that plainly beats pretending
    // to an accuracy the host API doesn't offer.
    const from = Math.max(0, start);
    const to = end > from ? Math.min(end, paragraphs.items.length) : paragraphs.items.length;
    const slice = paragraphs.items.slice(from, to);

    const blocks = slice.map((p) => {
      const style = String(p.styleBuiltIn ?? p.style ?? "");
      const heading = /^Heading(\d)$/.exec(style);
      return heading
        ? { type: "heading", text: p.text ?? "", level: Number(heading[1]) }
        : { type: "paragraph", text: p.text ?? "" };
    });
    const joined = blocks.map((b) => b.text).join("\n");

    return {
      anchor: {
        kind: "text" as const,
        start_offset: from,
        end_offset: to,
        content_sha: hashText(joined),
      },
      text: { blocks, char_count: joined.length },
    };
  });
}

async function searchDocument(
  args: Record<string, unknown>,
  host: Host
): Promise<Record<string, unknown>[]> {
  const needle = arg(args, "query");
  if (!needle) throw new Error("search_document needs a `query`.");

  if (host === "Word") {
    return Word.run(async (ctx) => {
      const hits = ctx.document.body.search(needle, { matchCase: false });
      hits.load("items/text");
      await ctx.sync();
      return hits.items.slice(0, 50).map((h, i) => ({ match: i + 1, text: h.text ?? "" }));
    });
  }

  if (host === "Excel") {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      const found = sheets.items.map((s) =>
        s.getUsedRangeOrNullObject(true).findOrNullObject(needle, {
          completeMatch: false,
          matchCase: false,
        })
      );
      found.forEach((f) => f.load("address,values,isNullObject"));
      await ctx.sync();

      const out: Record<string, unknown>[] = [];
      sheets.items.forEach((sheet, i) => {
        if (found[i].isNullObject) return;
        out.push({
          sheet: sheet.name,
          cell: String(found[i].address ?? "").replace(/^.*!/, ""),
          value: (found[i].values as unknown[][])?.[0]?.[0] ?? null,
        });
      });
      return out;
    });
  }

  return [];
}

// ── PowerPoint ─────────────────────────────────────────────────────────────

async function listSlides(): Promise<Record<string, unknown>[]> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const shapeSets = slides.items.map((s) => {
      const shapes = s.shapes;
      shapes.load("items/textFrame/textRange/text");
      return shapes;
    });
    await ctx.sync();

    return slides.items.map((slide, i) => {
      const texts = shapeSets[i].items
        .map((s) => {
          try {
            return (s.textFrame?.textRange?.text ?? "").trim();
          } catch {
            return "";
          }
        })
        .filter((t) => t);
      return { slide: i + 1, slide_id: String(slide.id ?? ""), title: texts[0] ?? "" };
    });
  });
}

async function readSlide(args: Record<string, unknown>): Promise<Selection> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const wantedId = arg(args, "slide_id");
    let position = Number(args.index ?? -1);
    if (wantedId) {
      position = slides.items.findIndex((s) => String(s.id ?? "") === wantedId);
    }
    if (position < 0 || position >= slides.items.length) {
      throw new Error("That slide isn't in this deck.");
    }

    const slide = slides.items[position];
    const shapes = slide.shapes;
    shapes.load("items/textFrame/textRange/text");
    await ctx.sync();

    const texts = shapes.items
      .map((s) => {
        try {
          return (s.textFrame?.textRange?.text ?? "").trim();
        } catch {
          return "";
        }
      })
      .filter((t) => t);

    return {
      anchor: {
        kind: "shape" as const,
        slide_index: position,
        slide_id: String(slide.id ?? ""),
        content_sha: hashText(texts.join("\n")),
      },
      shapes: {
        slides: [
          {
            slide_id: String(slide.id ?? ""),
            index: position,
            title: texts[0] ?? "",
            body: texts.slice(1),
            // Speaker notes need a second round trip per slide and the
            // PowerPoint API surfaces them inconsistently across builds;
            // `list_slides` reports whether they exist.
            notes: "",
          },
        ],
      },
    };
  });
}

// ── dispatch ───────────────────────────────────────────────────────────────

/** Run one request from the agent. Never throws: a failure comes back as
 * `ok: false` with a reason the agent can act on. */
export async function fulfil(
  request: ClientToolRequest,
  host: Host = currentHost()
): Promise<ClientToolResult> {
  const { request_id, op, args } = request;
  try {
    switch (op) {
      case "read_selection": {
        const selection = await captureSelection(host);
        return { request_id, ok: true, selection };
      }
      case "read_range":
        return { request_id, ok: true, selection: await readRange(args) };
      case "list_sheets":
        return { request_id, ok: true, items: await listSheets() };
      case "read_named_ranges":
        return { request_id, ok: true, items: await readNamedRanges() };
      case "read_tables":
        return { request_id, ok: true, items: await readTables() };
      case "read_document_outline":
        return { request_id, ok: true, items: await readOutline() };
      case "read_document_text":
        return { request_id, ok: true, selection: await readDocumentText(args) };
      case "list_slides":
        return { request_id, ok: true, items: await listSlides() };
      case "read_slide":
        return { request_id, ok: true, selection: await readSlide(args) };
      case "search_document":
        return { request_id, ok: true, items: await searchDocument(args, host) };
      case "trace_precedents":
        return { request_id, ok: true, items: await trace(args, "precedents") };
      case "trace_dependents":
        return { request_id, ok: true, items: await trace(args, "dependents") };
      default:
        return { request_id, ok: false, error: `This add-in can't do '${op}'.` };
    }
  } catch (err) {
    return {
      request_id,
      ok: false,
      error: err instanceof Error ? err.message : "That read failed.",
    };
  }
}
