/** The Office taskpane's client for `/api/v1/addons/*`.
 *
 * This replaces the SPA chat endpoints the taskpane used to call
 * (`/organizations/{org}/conversations`, `/chat/tables/extract`). Those work,
 * but they return markdown for a browser to render — and this add-in drives a
 * document API. `/addons/ask` returns the answer already split into typed
 * blocks with tables lifted out, plus the insertion plan and the citations,
 * in one round trip instead of four.
 *
 * Moving over is also what makes document context possible at all: `context`
 * only exists on the add-in surface.
 *
 * Header conventions (platform, user clock) match `chatClient.ts` so
 * server-side "today" and timezone-anchored answers behave identically on both.
 */

import { SelectionAnchor } from "./document/anchor";
import { Selection } from "./document/selection";

// ── wire shapes (mirrors app/api/v1/addons/schemas.py) ──────────────────────

export interface ContextPolicy {
  enabled: boolean;
  scope_max: "none" | "selection" | "sheet" | "document";
  max_cells: number;
  allow_formulas: boolean;
  allow_client_tools: boolean;
  allow_index: boolean;
  journal_max_cells: number;
  allow_writes: boolean;
  allow_unattended: boolean;
  default_permission_mode: "manual" | "auto" | "unattended";
}

export interface Bootstrap {
  user_id: string;
  email: string;
  full_name: string | null;
  organizations: { id: string; name: string; is_admin: boolean }[];
  branding: { app_name: string; logo_url: string };
  context_policy: ContextPolicy;
}

export interface AddonAgent {
  id: string;
  name: string;
  description: string;
}

export interface DocumentRef {
  host: string;
  document_id: string;
  title?: string;
  url?: string | null;
}

export interface Citation {
  anchor: SelectionAnchor;
  label: string;
}

export interface CitedBlock {
  type: string;
  text: string;
  level: number;
  table_index: number | null;
  citations: number[];
}

export interface XlsxTable {
  name: string;
  columns: string[];
  rows: (string | number | null)[][];
  chart_type: string | null;
}

export interface AddonAnswer {
  conversation_id: string;
  surface: string;
  prose: string;
  markdown?: string;
  blocks: CitedBlock[];
  tables: XlsxTable[];
  charts: { table_index: number; chart_type: string; native: boolean; title: string }[];
  intent: {
    artifact: string;
    requested_surface: string | null;
    surface_matches: boolean;
    confirm_prompt: string;
    is_update?: boolean;
  };
  plan: {
    operations: unknown[];
    summary: string;
    steps: string[];
    requires_confirmation: boolean;
  };
  notes: string[];
  citations: Citation[];
  context_snapshot_id: string;
  context_anchor: SelectionAnchor | null;
}

export interface DocumentContext {
  document: DocumentRef;
  scope: "none" | "selection" | "sheet" | "document";
  selection?: Selection | null;
  snapshot_id?: string;
  index_revision?: string;
  capabilities: string[];
  client_version?: string;
  permission_mode?: "manual" | "auto" | "unattended";
}

export class AddonApiError extends Error {
  constructor(
    public status: number,
    message: string,
    public detail?: unknown
  ) {
    super(message);
  }
}

/** The server no longer holds a snapshot id we claimed. A normal branch: the
 * caller re-sends the selection and retries once. */
export class ContextRequiredError extends AddonApiError {
  constructor(public snapshotId: string) {
    super(409, "The selection needs to be re-sent.");
  }
}

import { request } from "./api";

// ── endpoints ───────────────────────────────────────────────────────────────

/** Who am I, which orgs, the branding, and what this deployment lets us share
 * from the open file — one call instead of three. */
export function bootstrap(): Promise<Bootstrap> {
  return request("/addons/bootstrap");
}

export function listAgents(orgId: string): Promise<AddonAgent[]> {
  return request(`/addons/organizations/${orgId}/agents`);
}

/** Upload a selection and get the id to re-use on later turns. */
export function putSnapshot(
  orgId: string,
  document: DocumentRef,
  _scope: DocumentContext["scope"],
  selection: Selection
): Promise<{ context_id: string; notes: string[] }> {
  const payload = {
    host: document.host,
    document_id: document.document_id,
    kind: selection.anchor.kind,
    grid: selection.grid ? {
      sheet: selection.anchor.sheet_name || "",
      range: selection.anchor.a1_range || "",
      headers: [],
      rows: selection.grid.rows
    } : undefined,
    blocks: selection.text?.blocks?.map(b => ({
      text: b.text,
      heading_path: []
    })) || []
  };

  return request(`/addons/organizations/${orgId}/document-context`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

/** Upload the structural map of the open file.
 *
 * Structure only — sheet names, used ranges, headers, the heading tree. This
 * is what makes "it can see my whole workbook" affordable: the agent gets a
 * map in every prompt and pays for cell values only when it reads them.
 */
export function putIndex(
  orgId: string,
  index: unknown
): Promise<{ revision: string; sheet_count: number; approx_tokens: number; notes: string[] }> {
  return request(`/addons/organizations/${orgId}/context/index`, {
    method: "POST",
    body: JSON.stringify(index),
  });
}

export function ask(
  orgId: string,
  body: {
    text: string;
    surface: string;
    conversation_id?: string | null;
    agent_id?: string | null;
    context?: DocumentContext | null;
  }
): Promise<AddonAnswer> {
  const payload = {
    ...body,
    context_id: body.context?.snapshot_id || null,
  };
  return request(`/addons/organizations/${orgId}/ask`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

/** Ask, and recover once from a stale snapshot id.
 *
 * The cache is allowed to be wrong — that is the whole point of `409
 * context_required`. On a miss we upload the selection we still have in hand
 * and retry, so the user sees one answer and never the round trip.
 */
export async function askWithContext(
  orgId: string,
  body: Parameters<typeof ask>[1],
  onSnapshotStored?: (snapshotId: string) => void,
  selectionForRetry?: Selection | null
): Promise<AddonAnswer> {
  try {
    return await ask(orgId, body);
  } catch (err) {
    if (!(err instanceof ContextRequiredError) || !body.context || !selectionForRetry) {
      throw err;
    }
    const stored = await putSnapshot(
      orgId,
      body.context.document,
      body.context.scope,
      selectionForRetry
    );
    onSnapshotStored?.(stored.context_id);
    return ask(orgId, {
      ...body,
      context: { ...body.context, snapshot_id: stored.context_id, selection: null },
    });
  }
}

// ── background asks ─────────────────────────────────────────────────────────
//
// The synchronous `/ask` can't suspend, so the agent is never offered the
// document READ tools on it. `/ask/start` can — which is the only way it gets
// to look at another sheet mid-answer.

export interface JobRequest {
  request_id: string;
  tool: string;
  op: string;
  args: Record<string, unknown>;
  human_label: string;
  document_id: string;
}

export interface JobState {
  status: "running" | "needs_context" | "done" | "error";
  answer?: AddonAnswer;
  error?: string;
  requests?: JobRequest[];
  /** Human-readable, backend-sent progress label — e.g. "Search transactions…" —
   * while `status` is "running". */
  activity?: string;
}

export function askStart(
  orgId: string,
  body: Parameters<typeof ask>[1]
): Promise<{ job_id: string; status: string }> {
  const payload = {
    ...body,
    context_id: body.context?.snapshot_id || null,
  };
  return request(`/addons/organizations/${orgId}/ask/start`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export function getJob(jobId: string): Promise<JobState> {
  return request(`/addons/jobs/${jobId}`);
}

/** Answer the reads a suspended turn is waiting on. */
export function postJobContext(
  jobId: string,
  results: unknown[]
): Promise<{ accepted: number; status: string }> {
  return request(`/addons/jobs/${jobId}/context`, {
    method: "POST",
    body: JSON.stringify({ results }),
  });
}

/** The taskpane closed. Best-effort — the job's own TTL is the real backstop,
 * so a missed cancel costs a few pointless reads, not correctness. */
export function cancelJob(jobId: string): Promise<unknown> {
  return request(`/addons/jobs/${jobId}/cancel`, { method: "POST" });
}

/** Renders one table as a PNG, for hosts with no native chart object. */
export function renderChartPng(
  orgId: string,
  chartType: string,
  table: XlsxTable
): Promise<Blob> {
  return request(`/addons/organizations/${orgId}/charts/render`, {
    method: "POST",
    body: JSON.stringify({ chart_type: chartType, table }),
  });
}
