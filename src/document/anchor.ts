/** Where a piece of the document is, and whether it has changed since we read it.
 *
 * `SelectionAnchor` is the backend's one way of pointing at a region — it is
 * what a citation chip selects, and what an anchored edit targets.
 *
 * ## content_sha
 *
 * The server stores this and echoes it back; it never recomputes it. Its only
 * job is to let US notice that a region changed between reading it and writing
 * to it. That means the algorithm is our choice — the single property that
 * matters is that the same region hashes the same way on both sides of the
 * round trip.
 *
 * So: a cheap, stable encoding of the values, and the SAME function used at
 * capture time and at re-check time. Do not be tempted to "improve" it on one
 * path only.
 */

export interface SelectionAnchor {
  kind: "grid" | "text" | "shape" | "whole";
  sheet_name?: string;
  a1_range?: string;
  table_name?: string;
  start_offset?: number | null;
  end_offset?: number | null;
  named_range_id?: string;
  slide_id?: string;
  slide_index?: number | null;
  shape_ids?: string[];
  content_sha?: string;
}

/** FNV-1a, 32-bit, hex. Not cryptographic and doesn't need to be — this is a
 * change detector, not a signature. Cheap enough to run over a big range
 * without a visible pause, which a real digest in JS would not be. */
export function fingerprint(value: unknown): string {
  const text = JSON.stringify(value ?? null);
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    // hash * 16777619, kept in 32-bit range without overflowing to float
    hash = (hash + ((hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24))) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

/** The hash of a grid region's values. Row-major, `null` for empty, so a
 * blank cell and a missing cell agree. */
export function hashValues(values: unknown[][]): string {
  return fingerprint(values.map((row) => row.map((cell) => (cell === "" ? null : cell))));
}

export function hashText(text: string): string {
  return fingerprint(text);
}

/** Strip the sheet name off "Sheet1!B2:F87" — Office hands back addresses in
 * that form, the anchor keeps the two apart. */
export function splitAddress(address: string): { sheet: string; range: string } {
  const bang = address.lastIndexOf("!");
  if (bang < 0) return { sheet: "", range: address };
  return { sheet: address.slice(0, bang), range: address.slice(bang + 1) };
}

export function gridAnchor(address: string, values: unknown[][]): SelectionAnchor {
  const { sheet, range } = splitAddress(address);
  return {
    kind: "grid",
    sheet_name: sheet,
    a1_range: range,
    content_sha: hashValues(values),
  };
}

export function textAnchor(start: number, end: number, text: string): SelectionAnchor {
  return {
    kind: "text",
    start_offset: start,
    end_offset: end,
    content_sha: hashText(text),
  };
}

export function slideAnchor(index: number, id: string, text: string): SelectionAnchor {
  return {
    kind: "shape",
    slide_index: index,
    slide_id: id,
    content_sha: hashText(text),
  };
}

/** What the user sees on a citation chip before they click it. Mirrors
 * `SelectionAnchor.describe()` server-side so the two never disagree in the UI. */
export function describeAnchor(anchor: SelectionAnchor): string {
  if (anchor.kind === "grid") {
    const where = anchor.a1_range || anchor.table_name || "the selected cells";
    return anchor.sheet_name && !where.includes("!")
      ? `${anchor.sheet_name}!${where}`
      : where;
  }
  if (anchor.kind === "shape") {
    if (anchor.slide_index != null) return `slide ${anchor.slide_index + 1}`;
    return anchor.slide_id || "the selected shapes";
  }
  if (anchor.kind === "whole") return "the whole document";
  if (anchor.start_offset != null && anchor.end_offset != null) {
    return `characters ${anchor.start_offset}–${anchor.end_offset}`;
  }
  return "the selected text";
}
