/** Thin REST client for the taskpane — non-streaming only (dialogs/taskpanes
 * are a poor fit for SSE token-by-token rendering; the plan explicitly scopes
 * streaming out of v1). Mirrors the header/auth conventions of
 * `noah-frontend-v2/src/api/chat.ts` + `client.ts` + `lib/timezone.ts` so
 * server-side "today"/timezone-anchored answers behave the same here as on
 * the web app. */

// Imports removed
import { request, timeHeaders } from "./api";
import { API_BASE_URL, CLIENT_PLATFORM } from "./config";
import { getAccessToken } from "./tokenStorage";

export interface OrganizationMembership {
  organization_id: string;
  organization_name: string;
}

export interface CurrentUser {
  id: string;
  email: string;
  full_name: string | null;
  organizations: OrganizationMembership[];
}

export function getMe(): Promise<CurrentUser> {
  return request(`/auth/me`);
}

export interface Agent {
  id: string;
  name: string;
}

export interface Conversation {
  id: string;
  agent_id: string;
  title: string | null;
  create_time: string;
  update_time: string;
}

export interface XlsxTable {
  name: string;
  columns: string[];
  rows: (string | number | null)[][];
  chart_type: string | null;
}

export interface MessageContentBlock {
  type: string;
  text?: string;
  [key: string]: unknown;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "tool";
  content: MessageContentBlock[];
}

export function listAgents(orgId: string): Promise<Agent[]> {
  return request(`/organizations/${orgId}/agents`);
}

export function createConversation(
  orgId: string,
  agentId: string
): Promise<Conversation> {
  return request(`/organizations/${orgId}/conversations`, {
    method: "POST",
    body: JSON.stringify({ agent_id: agentId }),
  });
}

export function listConversations(
  orgId: string,
  limit: number = 50
): Promise<Conversation[]> {
  return request(`/organizations/${orgId}/conversations?limit=${limit}`);
}

export function listMessages(
  orgId: string,
  conversationId: string
): Promise<ChatMessage[]> {
  return request(`/organizations/${orgId}/conversations/${conversationId}/messages`);
}

export async function sendMessage(
  orgId: string,
  conversationId: string,
  text: string
): Promise<ChatMessage[]> {
  const res = await request<{ messages: ChatMessage[] }>(
    `/organizations/${orgId}/conversations/${conversationId}/messages`,
    { method: "POST", body: JSON.stringify({ text }) }
  );
  return res.messages;
}

export type ChatStreamEvent =
  | { type: "user_message"; data: any }
  | { type: "assistant_start"; data: Record<string, never> }
  | { type: "text_delta"; data: { index: number; text: string } }
  | { type: "tool_use_start"; data: { tool_use_id: string; name: string } }
  | { type: "tool_use_input_delta"; data: { tool_use_id: string; partial_json: string } }
  | { type: "tool_use_complete"; data: { tool_use_id: string; input: any } }
  | { type: "tool_executing"; data: { tool_use_id: string; name: string } }
  | { type: "tool_result"; data: any }
  | { type: "assistant_complete"; data: any }
  | { type: "assistant_attachments"; data: any }
  | { type: "export_status"; data: any }
  | { type: "error"; data: { message: string } }
  | { type: "done"; data: { chat_log_id?: string } };

export async function sendMessageStream(
  orgId: string,
  convId: string,
  text: string,
  onEvent: (e: ChatStreamEvent) => void,
  signal?: AbortSignal,
  contextBody?: any
): Promise<void> {
  const url = `${API_BASE_URL}/organizations/${orgId}/conversations/${convId}/messages/stream`;
  const token = await getAccessToken();
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Client-Platform": CLIENT_PLATFORM,
      Accept: "text/event-stream",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...timeHeaders(),
    },
    body: JSON.stringify({
      text,
      ...contextBody, // Inject Add-on context (snapshot_id, etc) if provided
    }),
    signal,
  });
  
  if (!res.ok || !res.body) {
    let detail = `HTTP ${res.status}`;
    try {
      const j = await res.json();
      if (j?.detail) detail = String(j.detail);
    } catch {}
    throw new Error(detail);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = "";
  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    let sep: number;
    while ((sep = buf.indexOf("\n\n")) !== -1) {
      const raw = buf.slice(0, sep);
      buf = buf.slice(sep + 2);
      let eventName = "message";
      const dataLines: string[] = [];
      for (const line of raw.split("\n")) {
        if (line.startsWith("event:")) eventName = line.slice(6).trim();
        else if (line.startsWith("data:")) dataLines.push(line.slice(5).trim());
      }
      if (dataLines.length === 0) continue;
      const dataStr = dataLines.join("\n");
      let data: unknown = dataStr;
      try {
        data = JSON.parse(dataStr);
      } catch {}
      onEvent({ type: eventName as any, data: data as any });
    }
  }
}

/** The assistant's plain-text answer, concatenated from its text blocks —
 * what gets inserted as prose (Word/PowerPoint/Docs/Slides). */
export function answerText(messages: ChatMessage[]): string {
  const assistant = messages.filter((m) => m.role === "assistant");
  return assistant
    .flatMap((m) => m.content)
    .filter((b) => b.type === "text" && typeof b.text === "string")
    .map((b) => b.text)
    .join("\n\n");
}

export function extractTables(
  orgId: string,
  markdown: string
): Promise<{ tables: XlsxTable[] }> {
  return request(`/organizations/${orgId}/chat/tables/extract`, {
    method: "POST",
    body: JSON.stringify({ markdown }),
  });
}

/** Renders one table + chart type to a PNG — used by hosts with no native
 * chart-object API (Word, PowerPoint). Excel builds its own native chart
 * client-side instead and never calls this. */
export async function renderChartPng(
  orgId: string,
  chartType: string,
  table: XlsxTable
): Promise<Blob> {
  return request(`/organizations/${orgId}/chat/charts/render`, {
    method: "POST",
    body: JSON.stringify({ chart_type: chartType, table }),
  });
}
