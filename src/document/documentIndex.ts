/** The structural map of the open file — what makes "it can see my workbook" true.
 *
 * This is STRUCTURE, never content: sheet names, used ranges, the header row,
 * named ranges, the heading tree, slide titles. A couple of thousand tokens
 * for a large file, and it sits in the agent's prompt on every turn.
 *
 * That distinction is the whole economics of the feature. Sending the workbook
 * would be a context overflow and a data-egress problem; sending the *map*
 * lets the agent answer "which sheet has the revenue?" with no reads at all,
 * and target the reads it does make. It pays for cell values only when it
 * decides it needs them.
 *
 * ## When we rebuild
 *
 * Lazily, on a dirty flag — never on selection. `DocumentSelectionChanged`
 * fires constantly and selecting a cell is not a structural change; rebuilding
 * there would make every click cost a workbook walk. We bump a counter when
 * the host tells us something actually changed, and rebuild at the next ask
 * if the counter has moved past what we last uploaded.
 *
 * `revision` is ours — a monotonic number, not a Drive revision id. Office has
 * none, and the server only ever compares it to the value it last stored.
 */

import { Host, currentHost } from "./capabilities";

export interface SheetIndex {
  name: string;
  used_range: string;
  row_count: number;
  column_count: number;
  header: string[];
  has_formulas: boolean;
  named_ranges: string[];
  tables: string[];
  hidden: boolean;
}

export interface OutlineEntry {
  text: string;
  level: number;
  start_offset?: number | null;
}

export interface SlideIndex {
  index: number;
  slide_id: string;
  title: string;
  has_notes: boolean;
}

export interface DocumentIndex {
  document: { host: string; document_id: string; title?: string };
  sheets: SheetIndex[];
  outline: OutlineEntry[];
  slides: SlideIndex[];
  revision: string;
  built_at: string;
  truncated: boolean;
}

/** Sheets we'll walk in one build. A workbook past this is a map nobody can
 * read and a prompt nobody should pay for — the server caps again at its own
 * limit and marks the result partial either way. */
const MAX_SHEETS = 60;
/** Header cells kept per sheet. */
const MAX_HEADER = 24;
const MAX_OUTLINE = 120;
const MAX_SLIDES = 120;

// ── dirty tracking ─────────────────────────────────────────────────────────

let revision = 0;
let uploadedRevision = "";
let handlersBound = false;

/** Something structural changed. Cheap — just moves a number; the rebuild
 * happens at the next ask. */
export function markDirty(): void {
  revision += 1;
}

export function isStale(): boolean {
  return String(revision) !== uploadedRevision;
}

export function markUploaded(rev: string): void {
  uploadedRevision = rev;
}

export function currentRevision(): string {
  return String(revision);
}

/** Ask the host to tell us when the document changes.
 *
 * Best-effort by design: the events are version-gated, and an add-in that
 * can't hear them still works — it just rebuilds only when the user asks it
 * to. Failing the session over a missing event would be the wrong trade.
 */
export async function watchForChanges(host: Host = currentHost()): Promise<void> {
  if (handlersBound) return;
  handlersBound = true;
  try {
    if (host === "Excel") {
      await Excel.run(async (ctx) => {
        // The collection-level event covers every sheet including ones added
        // later; per-sheet handlers would miss those.
        ctx.workbook.worksheets.onChanged.add(async () => markDirty());
        ctx.workbook.worksheets.onAdded.add(async () => markDirty());
        ctx.workbook.worksheets.onDeleted.add(async () => markDirty());
        await ctx.sync();
      });
    } else if (host === "Word") {
      await Word.run(async (ctx) => {
        ctx.document.onParagraphChanged.add(async () => markDirty());
        await ctx.sync();
      });
    }
    // PowerPoint has no change event; its index is rebuilt when the user asks
    // for a refresh, or on the first ask of a session.
  } catch {
    // Older host, or the event isn't supported. The map still builds on
    // demand — it just won't notice edits on its own.
    handlersBound = false;
  }
}

// ── builders ───────────────────────────────────────────────────────────────

async function buildExcel(): Promise<{ sheets: SheetIndex[]; truncated: boolean }> {
  return Excel.run(async (ctx) => {
    const worksheets = ctx.workbook.worksheets;
    worksheets.load("items/name,items/visibility");
    const workbookNames = ctx.workbook.names;
    workbookNames.load("items/name,items/value");
    await ctx.sync();

    const all = worksheets.items;
    const truncated = all.length > MAX_SHEETS;
    const items = all.slice(0, MAX_SHEETS);

    // `getUsedRangeOrNullObject` rather than `getUsedRange`: an empty sheet
    // throws on the latter, and one blank tab shouldn't fail the whole map.
    const used = items.map((sheet) => {
      const range = sheet.getUsedRangeOrNullObject(true);
      range.load("address,rowCount,columnCount,isNullObject");
      return range;
    });
    const tables = items.map((sheet) => {
      const collection = sheet.tables;
      collection.load("items/name");
      return collection;
    });
    await ctx.sync();

    // Second pass for header rows — only for sheets that have content, and
    // only the first row. This is the one part of the map that touches values.
    const headers = used.map((range) =>
      range.isNullObject ? null : range.getRow(0).getUsedRangeOrNullObject(true)
    );
    headers.forEach((row) => row?.load("values,isNullObject"));
    // Formula presence: a cheap boolean per sheet, so the agent knows whether
    // "why is this wrong" is even a formula question here.
    const formulaProbe = used.map((range) =>
      range.isNullObject ? null : range
    );
    formulaProbe.forEach((range) => range?.load("formulas"));
    await ctx.sync();

    const namesBySheet = new Map<string, string[]>();
    for (const named of workbookNames.items) {
      const value = String(named.value ?? "");
      const bang = value.indexOf("!");
      const sheet = bang > 0 ? value.slice(0, bang).replace(/^=/, "").replace(/'/g, "") : "";
      if (!namesBySheet.has(sheet)) namesBySheet.set(sheet, []);
      namesBySheet.get(sheet)!.push(named.name);
    }

    const sheets: SheetIndex[] = items.map((sheet, i) => {
      const range = used[i];
      const empty = range.isNullObject;
      const headerRow = headers[i];
      const headerValues =
        headerRow && !headerRow.isNullObject
          ? ((headerRow.values?.[0] ?? []) as unknown[])
              .slice(0, MAX_HEADER)
              .map((cell) => String(cell ?? "").trim())
              .filter((cell) => cell)
          : [];
      const formulas = (formulaProbe[i]?.formulas ?? []) as string[][];
      return {
        name: sheet.name,
        used_range: empty ? "" : String(range.address ?? "").replace(/^.*!/, ""),
        row_count: empty ? 0 : Number(range.rowCount ?? 0),
        column_count: empty ? 0 : Number(range.columnCount ?? 0),
        header: headerValues,
        has_formulas: formulas.some((row) =>
          row.some((cell) => typeof cell === "string" && cell.startsWith("="))
        ),
        named_ranges: namesBySheet.get(sheet.name) ?? [],
        tables: tables[i].items.map((t) => t.name),
        hidden: sheet.visibility !== Excel.SheetVisibility.visible,
      };
    });

    return { sheets, truncated };
  });
}

async function buildWord(): Promise<{ outline: OutlineEntry[]; truncated: boolean }> {
  return Word.run(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load("items/text,items/styleBuiltIn,items/style");
    await ctx.sync();

    const outline: OutlineEntry[] = [];
    for (const p of paragraphs.items) {
      const style = String(p.styleBuiltIn ?? p.style ?? "");
      const heading = /^Heading(\d)$/.exec(style);
      if (!heading) continue;
      const text = (p.text ?? "").trim();
      if (!text) continue;
      outline.push({ text, level: Number(heading[1]) });
      if (outline.length >= MAX_OUTLINE) return { outline, truncated: true };
    }
    return { outline, truncated: false };
  });
}

async function buildPowerPoint(): Promise<{ slides: SlideIndex[]; truncated: boolean }> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();

    const truncated = slides.items.length > MAX_SLIDES;
    const items = slides.items.slice(0, MAX_SLIDES);

    const shapeSets = items.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/name,items/textFrame/textRange/text");
      return shapes;
    });
    await ctx.sync();

    const out: SlideIndex[] = items.map((slide, i) => {
      const texts = shapeSets[i].items
        .map((s) => {
          try {
            return (s.textFrame?.textRange?.text ?? "").trim();
          } catch {
            // A picture or a line has no text frame and throws rather than
            // returning empty.
            return "";
          }
        })
        .filter((t) => t);
      return {
        index: i,
        slide_id: String(slide.id ?? ""),
        title: texts[0] ?? "",
        // Reading notes needs a second round trip per slide; the map only has
        // to say whether there ARE any, and the agent can read one if it cares.
        has_notes: false,
      };
    });

    return { slides: out, truncated };
  });
}

/** Build the map for the open file, or null when this host has none to give. */
export async function buildIndex(
  document: { host: string; document_id: string; title?: string },
  host: Host = currentHost()
): Promise<DocumentIndex | null> {
  const base: DocumentIndex = {
    document,
    sheets: [],
    outline: [],
    slides: [],
    revision: currentRevision(),
    built_at: new Date().toISOString(),
    truncated: false,
  };

  try {
    if (host === "Excel") {
      const { sheets, truncated } = await buildExcel();
      return { ...base, sheets, truncated };
    }
    if (host === "Word") {
      const { outline, truncated } = await buildWord();
      return { ...base, outline, truncated };
    }
    if (host === "PowerPoint") {
      const { slides, truncated } = await buildPowerPoint();
      return { ...base, slides, truncated };
    }
    return null;
  } catch (err) {
    // A map we couldn't build is a smaller answer, not a failed question.
    console.warn("Noah: could not build the document map", err);
    return null;
  }
}
