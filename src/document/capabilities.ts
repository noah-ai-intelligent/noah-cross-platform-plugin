/** What THIS build, in THIS host, can actually be asked to do.
 *
 * The backend builds the agent's tool schemas from this list
 * (`app/api/v1/addons/tools.py`), intersected with the ops it knows. That makes
 * the list a real contract, and both failure directions are quiet:
 *
 *   under-declare → the agent is never offered the op and silently works
 *                   around its absence
 *   over-declare  → the agent calls something this host can't do, the request
 *                   fails, and the turn degrades for no reason
 *
 * So it is PROBED, never hardcoded. The same add-in runs against Excel builds
 * years apart, and Office.js gates capability behind requirement sets —
 * `getDirectPrecedents` simply does not exist below ExcelApi 1.15.
 *
 * Write ops are deliberately absent for now: declaring none is how the
 * read-only build ships first, with no server flag involved (delivery plan D4).
 * Adding them here is what turns the action loop on at C5.
 */

export type Capability =
  | "read_selection"
  | "read_range"
  | "list_sheets"
  | "read_named_ranges"
  | "read_tables"
  | "read_document_outline"
  | "read_document_text"
  | "list_slides"
  | "read_slide"
  | "search_document"
  | "trace_precedents"
  | "trace_dependents";

export type Host = "Word" | "Excel" | "PowerPoint" | "GoogleSheets" | "GoogleDocs" | "GoogleSlides" | "Unknown";

export function currentHost(): Host {
  // @ts-ignore
  if (typeof window !== "undefined" && window.__GOOGLE_HOST__ && !window.__GOOGLE_HOST__.startsWith("<?")) {
    // @ts-ignore
    return window.__GOOGLE_HOST__ as Host;
  }
  // @ts-ignore
  if (typeof google !== "undefined" && google.script) {
    return "GoogleSheets"; // fallback
  }
  const h = window.Office?.context?.host;
  if (h === window.Office?.HostType.Word) return "Word";
  if (h === window.Office?.HostType.Excel) return "Excel";
  if (h === window.Office?.HostType.PowerPoint) return "PowerPoint";
  return "Unknown";
}

/** The `AddonHost` value the backend expects for this host. */
export function hostKey(host: Host = currentHost()): string {
  switch (host) {
    case "Word":
    case "GoogleDocs":
      return "docs";
    case "Excel":
    case "GoogleSheets":
      return "sheets";
    case "PowerPoint":
    case "GoogleSlides":
      return "slides";
    default:
      return "sheets";
  }
}

function supports(set: string, version: string): boolean {
  try {
    return Boolean(window.Office?.context?.requirements?.isSetSupported(set, version));
  } catch {
    // An older host can throw rather than return false. Treat any doubt as
    // "not supported" — under-declaring degrades, over-declaring breaks.
    return false;
  }
}

export function capabilities(host: Host = currentHost()): Capability[] {
  const caps: Capability[] = ["read_selection"];

  if (host === "Excel" || host === "GoogleSheets") {
    caps.push("read_range", "list_sheets");
    if (host === "GoogleSheets" || supports("ExcelApi", "1.4")) caps.push("read_named_ranges");
    if (host === "GoogleSheets" || supports("ExcelApi", "1.2")) caps.push("read_tables");
    // getDirectPrecedents / getDirectDependents. This is the whole of our
    // formula-tracing story: Apps Script has no equivalent, which is why
    // Sheets tracing is cut from v1 (delivery plan, D5).
    if (host === "Excel" && supports("ExcelApi", "1.15")) {
      caps.push("trace_precedents", "trace_dependents");
    }
  }

  if (host === "Word" || host === "GoogleDocs") {
    caps.push("read_document_outline", "read_document_text");
    if (host === "Word" && supports("WordApi", "1.3")) caps.push("search_document");
  }

  if (host === "PowerPoint" || host === "GoogleSlides") {
    caps.push("list_slides");
    // Reading one slide's body and speaker notes needs the richer shape API;
    // below it we can enumerate slides but not read into them.
    if (host === "PowerPoint" && supports("PowerPointApi", "1.3")) caps.push("read_slide");
  }

  return caps;
}

/** Logged with the session. When someone reports "it wouldn't trace the
 * formula", the first useful question is what this build actually declared —
 * the same source can behave differently on two Excel versions. */
export function capabilitySummary(host: Host = currentHost()): string {
  return `${hostKey(host)}: ${capabilities(host).join(",")}`;
}
