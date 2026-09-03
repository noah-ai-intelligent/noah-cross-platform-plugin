import React, { useMemo } from "react";
import type { AddonAnswer, Citation, XlsxTable, EditPlanOut } from "../../addonClient";
import type { ChatMessage } from "../../chatClient";
import { describeAnchor } from "../../document/anchor";
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import { ChartPreview } from './ChartPreview';

export function parseMarkdownTable(markdown: string): XlsxTable | null {
  if (!markdown) return null;
  const lines = markdown.split("\n").map((l) => l.trim()).filter(Boolean);
  const tableLines = lines.filter((l) => l.startsWith("|") && l.endsWith("|"));
  if (tableLines.length < 3) return null;

  const headerLine = tableLines[0];
  const columns = headerLine
    .split("|")
    .slice(1, -1)
    .map((c) => c.trim().replace(/^\*\*(.*)\*\*$/, "$1"));

  const separatorLine = tableLines[1];
  const isSeparator = /^\|(\s*:?-+:?\s*\|)+$/.test(separatorLine);
  const dataLines = isSeparator ? tableLines.slice(2) : tableLines.slice(1);

  const rows = dataLines.map((line) =>
    line
      .split("|")
      .slice(1, -1)
      .map((cell) => cell.trim().replace(/^\*\*(.*)\*\*$/, "$1"))
  );

  if (columns.length === 0 || rows.length === 0) return null;

  return {
    name: "NoahTable",
    columns,
    rows,
    chart_type: null,
  };
}

export type HistoryEntry = {
  id: string;
  question: string;
  preview: string;
  timestamp: number;
};

export function CellReference({ citation, onClick, disabled }: { citation: Citation; onClick?: () => void; disabled?: boolean }) {
  if (!citation || !citation.anchor) return null;
  return (
    <button
      className={`inline-flex items-center gap-1 min-h-[22px] px-2 rounded-full border border-zinc-200 bg-white shadow-xs text-[11px] font-medium transition-all ${disabled ? "opacity-50 cursor-not-allowed text-zinc-400" : "text-emerald-700 cursor-pointer hover:border-emerald-300 hover:shadow-sm hover:text-emerald-800"
        }`}
      onClick={disabled ? undefined : onClick}
      title={disabled ? "Feature not yet available in Slides" : describeAnchor(citation.anchor)}
      disabled={disabled}
    >
      <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
        <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
        <polyline points="15 3 21 3 21 9"></polyline>
        <line x1="10" y1="14" x2="21" y2="3"></line>
      </svg>
      <span>{describeAnchor(citation.anchor)}</span>
    </button>
  );
}

export function SelectedContext({
  enabled,
  onToggle,
  hint,
}: {
  enabled: boolean;
  onToggle: (v: boolean) => void;
  hint: string;
}) {
  if (!enabled || !hint) return null;
  return (
    <div className="flex items-center justify-between bg-emerald-50/70 border-b border-emerald-100 py-1.5 px-3 text-[12px] text-emerald-900 transition-all">
      <div className="flex items-center gap-1.5 font-medium truncate">
        <span className="relative flex h-2 w-2">
          <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
          <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
        </span>
        <span className="text-emerald-700">Reading:</span>
        <span className="font-mono text-[11px] bg-emerald-100/80 px-1.5 py-0.5 rounded text-emerald-800 font-semibold">{hint}</span>
      </div>

      <button
        className="inline-flex items-center justify-center border-none bg-transparent text-emerald-600 cursor-pointer p-0.5 hover:text-emerald-900 transition-colors rounded-full hover:bg-emerald-100/60"
        onClick={() => onToggle(false)}
        aria-label="Remove selection context"
        title="Detach selection"
      >
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <line x1="18" y1="6" x2="6" y2="18"></line>
          <line x1="6" y1="6" x2="18" y2="18"></line>
        </svg>
      </button>
    </div>
  );
}

export function ActionBar({ children }: { children: React.ReactNode }) {
  return <div className="flex items-center gap-1.5 flex-wrap mt-2 pt-1 border-t border-zinc-100">{children}</div>;
}

export function LoadingState({ activity }: { activity: string }) {
  return (
    <div className="flex flex-col gap-1 items-start my-1 animate-fadeIn">
      <div className="max-w-[85%] border border-zinc-200/80 bg-white shadow-xs rounded-2xl rounded-tl-xs px-3.5 py-2.5 flex items-center gap-1.5">
        <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-dot-1" />
        <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-dot-2" />
        <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-dot-3" />
        <span className="ml-1 text-xs text-zinc-500">{activity || "Thinking"}</span>
      </div>
    </div>
  );
}

export function ErrorState({ message }: { message: string }) {
  return (
    <div className="flex items-start gap-2 border border-red-200 bg-red-50/90 text-red-700 rounded-xl p-3 text-[13px] shadow-2xs animate-fadeIn my-1">
      <svg className="flex-none mt-0.5" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
        <circle cx="12" cy="12" r="10"></circle>
        <line x1="12" y1="8" x2="12" y2="12"></line>
        <line x1="12" y1="16" x2="12.01" y2="16"></line>
      </svg>
      <div className="flex-1">{message}</div>
    </div>
  );
}

export function UserMessage({ text }: { text: string }) {
  return (
    <div className="flex flex-col gap-1 items-end animate-fadeIn my-1">
      <div className="max-w-[85%] text-[13px] leading-relaxed whitespace-pre-wrap break-words bg-[#5FBD83] text-white rounded-2xl rounded-tr-xs py-2.5 px-3.5 shadow-sm font-normal">
        {text}
      </div>
    </div>
  );
}

export function PaginatedTable({ children, ...props }: any) {
  const [currentPage, setCurrentPage] = React.useState(1);
  const [pageSize, setPageSize] = React.useState(5);

  const childrenArray = React.Children.toArray(children);
  const thead = childrenArray.find((c: any) => c.type === "thead" || c.props?.node?.tagName === "thead");
  const tbody = childrenArray.find((c: any) => c.type === "tbody" || c.props?.node?.tagName === "tbody");

  if (!tbody) {
    return (
      <div className="overflow-x-auto rounded-lg">
        <table {...props}>{children}</table>
      </div>
    );
  }

  const rows = React.Children.toArray((tbody as any).props.children);
  const totalPages = Math.ceil(rows.length / pageSize);

  const startIndex = (currentPage - 1) * pageSize;
  const currentRows = rows.slice(startIndex, startIndex + pageSize);

  const paginatedTbody = React.cloneElement(tbody as React.ReactElement, {}, currentRows);

  return (
    <details className="flex flex-col gap-2 my-4 border border-zinc-200/80 rounded-xl overflow-hidden group shadow-xs">
      <summary className="px-3.5 py-2 text-[12px] font-semibold text-zinc-500 cursor-pointer select-none bg-zinc-50/80 hover:bg-zinc-100 flex items-center gap-2 transition-colors">
        <svg className="w-3.5 h-3.5 text-zinc-400 group-open:rotate-90 transition-transform duration-200" viewBox="0 0 16 16" fill="none" stroke="currentColor">
          <path d="M6 12L10 8L6 4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
        View Table Data
      </summary>
      <div className="overflow-x-auto border-t border-zinc-100">
        <table className="w-full text-left border-collapse" {...props}>
          {thead}
          {paginatedTbody}
        </table>
      </div>
      {rows.length > 5 && (
        <div className="flex items-center justify-between text-[11px] text-zinc-500 bg-zinc-50 px-2 py-1.5 rounded-md border border-zinc-100">
          <div className="flex items-center gap-1.5">
            <span>Show</span>
            <select
              className="border border-zinc-200 rounded px-1 py-0.5 bg-white text-zinc-700 outline-none focus:border-emerald-400"
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
            >
              <option value={5}>5</option>
              <option value={10}>10</option>
              <option value={20}>20</option>
              <option value={50}>50</option>
            </select>
            <span>entries</span>
          </div>
          <div className="flex items-center gap-1">
            <button
              disabled={currentPage === 1}
              onClick={() => setCurrentPage((p) => p - 1)}
              className="disabled:opacity-30 p-1 rounded hover:bg-zinc-200/60 transition-colors cursor-pointer disabled:cursor-not-allowed flex items-center justify-center text-zinc-600"
              title="Previous page"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="15 18 9 12 15 6"></polyline>
              </svg>
            </button>
            <span className="font-medium text-zinc-600 min-w-[2.5rem] text-center">
              {currentPage} / {totalPages || 1}
            </span>
            <button
              disabled={currentPage === totalPages || totalPages === 0}
              onClick={() => setCurrentPage((p) => p + 1)}
              className="disabled:opacity-30 p-1 rounded hover:bg-zinc-200/60 transition-colors cursor-pointer disabled:cursor-not-allowed flex items-center justify-center text-zinc-600"
              title="Next page"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="9 18 15 12 9 6"></polyline>
              </svg>
            </button>
          </div>
        </div>
      )}
    </details>
  );
}

export function ParsedMarkdown({ content }: { content: string }) {
  if (!content) return null;

  // Use fromCharCode to prevent Google Apps Script HtmlService from failing to parse literal '<think' in the bundle
  const lt = String.fromCharCode(60);
  const gt = String.fromCharCode(62);
  const regex = new RegExp(`(${lt}(think|thought|thinking)${gt}([\\s\\S]*?)(${lt}/\\2${gt}|$))`, 'gi');
  const parts = [];
  let lastIdx = 0;
  let m;

  while ((m = regex.exec(content)) !== null) {
    if (m.index > lastIdx) {
      parts.push({ type: 'text', content: content.substring(lastIdx, m.index) });
    }
    parts.push({ type: 'think', tag: m[2], content: m[3] });
    lastIdx = regex.lastIndex;
  }

  if (lastIdx < content.length) {
    const remaining = content.substring(lastIdx);
    if (remaining.trim()) {
      parts.push({ type: 'text', content: remaining });
    }
  }

  return (
    <>
      {parts.map((p, i) => {
        if (p.type === 'think') {
          // If it's the last part and doesn't have a closing tag, we assume it's streaming
          const isStreaming = i === parts.length - 1 && !content.toLowerCase().includes(`${lt}/${(p.tag || "").toLowerCase()}${gt}`, lastIdx - 10);
          return (
            <div key={i} className={`flex items-start gap-2 mb-2 ${isStreaming ? "animate-pulse" : "opacity-80"}`}>
              <svg className="w-4 h-4 text-zinc-400 mt-0.5 flex-shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="9 18 15 12 9 6" />
              </svg>
              <div className="text-[13px] text-zinc-500 font-medium font-serif italic">
                {p.content || "Thinking..."} {isStreaming ? "" : ">"}
              </div>
            </div>
          );
        }
        return (
          <ReactMarkdown
            key={i}
            remarkPlugins={[remarkGfm]}
            components={{
              table: PaginatedTable,
              code({ node, inline, className, children, ...props }: any) {
                const match = /language-(\w+)/.exec(className || "");
                if (!inline && match && match[1] === "chart") {
                  const jsonString = String(children).replace(/\n$/, "");
                  return <ChartPreview jsonString={jsonString} />;
                }
                return !inline ? (
                  <pre className="bg-zinc-100 p-3 rounded-lg overflow-x-auto text-xs my-2 font-mono">
                    <code className={className} {...props}>
                      {children}
                    </code>
                  </pre>
                ) : (
                  <code className="bg-zinc-100 px-1 py-0.5 rounded text-[0.9em] font-mono text-zinc-800" {...props}>
                    {children}
                  </code>
                );
              },
            }}
          >
            {p.content}
          </ReactMarkdown>
        );
      })}
    </>
  );
}


export function AssistantMessage({
  message,
  answer,
  host,
  onInsertProse,
  onInsertTable,
  onInsertReport,
  onCitation,
  onApplyEditPlan,
  onRejectEditPlan,
  requiresConfirmation = true,
}: {
  message: ChatMessage;
  answer?: AddonAnswer | null;
  host: string;
  onInsertProse: (text: string) => void;
  onInsertTable?: (table: XlsxTable) => void;
  onInsertReport?: (answer: AddonAnswer) => void;
  onCitation?: (c: Citation) => void;
  onApplyEditPlan?: (plan: EditPlanOut) => void;
  onRejectEditPlan?: (plan: EditPlanOut) => void;
  requiresConfirmation?: boolean;
}) {
  const textBlocks = message.content.filter((b) => b.type === "text" && typeof b.text === "string" && b.text.trim().length > 0);
  const fullText = textBlocks.map((b) => b.text).join("\n\n");
  const targetProse = fullText || answer?.prose || "";

  const tableFromAnswer = answer?.tables && answer.tables.length > 0 ? answer.tables[0] : null;
  const parsedTable = useMemo(() => {
    if (tableFromAnswer) return tableFromAnswer;
    return parseMarkdownTable(fullText || answer?.markdown || "");
  }, [tableFromAnswer, fullText, answer]);

  const hasValidAnswerBlocks =
    answer &&
    answer.blocks &&
    answer.blocks.length > 0 &&
    answer.blocks.some((b) => b.type === "table" || (b.text && b.text.trim().length > 0));

  if (!fullText.trim() && !hasValidAnswerBlocks && !answer?.markdown && !answer?.prose) {
    return null;
  }

  const hasReport = host === "Excel" && answer?.plan?.operations && answer.plan.operations.length > 0;
  const hasActions = host === "Excel" ? Boolean(hasReport || (parsedTable && onInsertTable)) : Boolean(targetProse || (parsedTable && onInsertTable));

  return (
    <div className="flex flex-col gap-1 items-start animate-fadeIn my-1">
      <div className="max-w-[92%] border border-zinc-200/80 bg-white shadow-xs rounded-2xl rounded-tl-xs p-3.5 transition-shadow hover:shadow-md">
        <div className="prose-noah break-words overflow-x-auto">
          {hasValidAnswerBlocks ? (
            answer.blocks.map((block, i) => {
              if (!block.text || !block.text.trim()) {
                if (block.citations && block.citations.length > 0) {
                  return (
                    <div key={i} className="mb-2 last:mb-0">
                      {block.citations.map((index) => {
                        const citation = answer.citations[index];
                        if (!citation) return null;
                        return <CellReference key={index} citation={citation} onClick={() => onCitation && onCitation(citation)} />;
                      })}
                    </div>
                  );
                }
                return null;
              }
              return (
                <div key={i} className="mb-2 last:mb-0">
                  <ParsedMarkdown content={block.text} />
                  {block.citations &&
                    block.citations.map((index) => {
                      const citation = answer.citations[index];
                      if (!citation) return null;
                      return <CellReference key={index} citation={citation} onClick={() => onCitation && onCitation(citation)} />;
                    })}
                </div>
              );
            })
          ) : (
            <ParsedMarkdown content={fullText || answer?.markdown || answer?.prose || ""} />
          )}
        </div>

        {answer?.edit_plan && (
          <div className="mt-3 border-t border-zinc-200/80 pt-3">
            <div className="text-[13px] font-semibold text-zinc-800 mb-1 flex items-center gap-1.5">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-emerald-600">
                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
              </svg>
              Proposed Edits
            </div>
            {answer.edit_plan.summary && (
              <p className="text-[12px] leading-relaxed text-zinc-600 mb-2">
                {answer.edit_plan.summary}
              </p>
            )}
            {answer.edit_plan.steps && answer.edit_plan.steps.length > 0 && (
              <ul className="list-disc list-inside text-[12px] leading-relaxed text-zinc-600 mb-3 space-y-1 ml-1">
                {answer.edit_plan.steps.map((step, i) => (
                  <li key={i}>{step}</li>
                ))}
              </ul>
            )}
            <div className="flex gap-2">
              <button
                className="inline-flex items-center justify-center px-3 py-1.5 border border-transparent rounded-md shadow-sm text-xs font-medium text-white bg-emerald-600 hover:bg-emerald-700 transition-colors cursor-pointer"
                onClick={() => onApplyEditPlan && onApplyEditPlan(answer.edit_plan!)}
              >
                Apply Edits
              </button>
              <button
                className="inline-flex items-center justify-center px-3 py-1.5 border border-zinc-300 rounded-md shadow-sm text-xs font-medium text-zinc-700 bg-white hover:bg-zinc-50 transition-colors cursor-pointer"
                onClick={() => onRejectEditPlan && onRejectEditPlan(answer.edit_plan!)}
              >
                Reject
              </button>
            </div>
          </div>
        )}


        {hasActions && (
          (answer ? (answer.plan?.requires_confirmation ?? requiresConfirmation) : true) ? (
            <ActionBar>
              {host === "Excel" ? (
                hasReport && onInsertReport ? (
                  <button
                    className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 border border-emerald-200 bg-emerald-50 rounded-md text-xs font-medium text-emerald-700 hover:bg-emerald-100 transition-colors cursor-pointer"
                    onClick={() => onInsertReport(answer!)}
                    title="Insert dashboard"
                  >
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M3 3h18v18H3z"></path>
                      <path d="M3 9h18"></path>
                      <path d="M9 21V9"></path>
                    </svg>
                    Insert Dashboard into sheet
                  </button>
                ) : parsedTable && onInsertTable ? (
                  <button
                    className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 border border-emerald-200 bg-emerald-50 rounded-md text-xs font-medium text-emerald-700 hover:bg-emerald-100 transition-colors cursor-pointer"
                    onClick={() => onInsertTable(parsedTable)}
                    title="Insert table"
                  >
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M12 5v14M5 12h14"></path>
                    </svg>
                    Insert Table into sheet
                  </button>
                ) : null
              ) : (
                <button
                  className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 border border-emerald-200 bg-emerald-50 rounded-md text-xs font-medium text-emerald-700 hover:bg-emerald-100 transition-colors cursor-pointer"
                  onClick={() => {
                    if (parsedTable && !targetProse.trim() && onInsertTable) {
                      onInsertTable(parsedTable);
                    } else if (onInsertProse && targetProse) {
                      onInsertProse(targetProse);
                    }
                  }}
                  title="Insert answer"
                >
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M12 5v14M5 12h14"></path>
                  </svg>
                  Insert into document
                </button>
              )}
            </ActionBar>
          ) : null
        )}
      </div>
    </div>
  );
}

export function ConversationView({
  messages,
  latestAnswer,
  busy,
  activity,
  error,
  host,
  onInsertProse,
  onInsertTable,
  onInsertReport,
  onCitation,
  onApplyEditPlan,
  onRejectEditPlan,
  bottomRef,
  requiresConfirmation = true,
}: {
  messages: ChatMessage[];
  latestAnswer?: AddonAnswer | null;
  busy: boolean;
  activity: string;
  error: string | null;
  host: string;
  onInsertProse: (text: string) => void;
  onInsertTable?: (table: XlsxTable) => void;
  onInsertReport?: (answer: AddonAnswer) => void;
  onCitation?: (c: Citation) => void;
  onApplyEditPlan?: (plan: EditPlanOut) => void;
  onRejectEditPlan?: (plan: EditPlanOut) => void;
  bottomRef?: React.RefObject<HTMLDivElement>;
  requiresConfirmation?: boolean;
}) {
  const displayMessages = useMemo(() => {
    const result: ChatMessage[] = [];
    let currentAssistantMsg: ChatMessage | null = null;

    for (const msg of messages) {
      if (msg.role === "user") {
        if (currentAssistantMsg) {
          result.push(currentAssistantMsg);
          currentAssistantMsg = null;
        }
        result.push(msg);
      } else if (msg.role === "assistant") {
        const validBlocks = (msg.content || []).filter(
          (b) => b.type === "text" && typeof b.text === "string" && b.text.trim().length > 0
        );
        if (validBlocks.length > 0) {
          if (!currentAssistantMsg) {
            currentAssistantMsg = {
              id: msg.id,
              role: "assistant",
              content: validBlocks.map((b) => ({ ...b })),
            };
          } else {
            currentAssistantMsg.content.push(...validBlocks.map((b) => ({ ...b })));
          }
        }
      }
    }
    if (currentAssistantMsg) {
      result.push(currentAssistantMsg);
    }
    return result;
  }, [messages]);

  return (
    <div className="flex flex-col gap-3 p-3.5 pb-4">
      {displayMessages.map((msg, index) => {
        const isLast = index === displayMessages.length - 1;
        const isLatest = isLast && latestAnswer && !busy;
        return msg.role === "user" ? (
          <UserMessage key={msg.id} text={msg.content[0]?.text || ""} />
        ) : (
          <AssistantMessage
            key={msg.id}
            message={msg}
            answer={isLatest ? latestAnswer : null}
            host={host}
            onInsertProse={onInsertProse}
            onInsertTable={onInsertTable}
            onInsertReport={onInsertReport}
            onCitation={onCitation}
            onApplyEditPlan={onApplyEditPlan}
            onRejectEditPlan={onRejectEditPlan}
            requiresConfirmation={requiresConfirmation}
          />
        );
      })}

      {busy && latestAnswer && (
        <AssistantMessage
          message={{ id: "streaming", role: "assistant", content: [] }}
          answer={latestAnswer}
          host={host}
          onInsertProse={onInsertProse}
          onInsertTable={onInsertTable}
          onInsertReport={onInsertReport}
          onCitation={onCitation}
          requiresConfirmation={requiresConfirmation}
        />
      )}

      {busy && !latestAnswer && <LoadingState activity={activity} />}
      {!busy && error && <ErrorState message={error} />}
      <div ref={bottomRef} className="h-px w-full" />
    </div>
  );
}

export function ConversationList({
  entries,
  onSelect,
}: {
  entries: HistoryEntry[];
  onSelect: (entry: HistoryEntry) => void;
}) {
  if (!entries.length) {
    return (
      <div className="flex flex-col items-center justify-center gap-2 h-full text-zinc-400 text-[13px] text-center py-10 px-5">
        <div className="w-10 h-10 rounded-full bg-zinc-100 flex items-center justify-center text-zinc-400 mb-1">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
          </svg>
        </div>
        <span className="font-medium text-zinc-700">No conversations yet</span>
        <span className="text-[12px] text-zinc-400">Questions you ask Noah will show up here.</span>
      </div>
    );
  }
  return (
    <div className="divide-y divide-zinc-100">
      {entries
        .slice()
        .reverse()
        .map((entry) => (
          <button
            key={entry.id}
            className="block w-full text-left border-none bg-transparent py-3 px-3 cursor-pointer rounded-lg hover:bg-emerald-50/50 transition-colors my-1"
            onClick={() => onSelect(entry)}
          >
            <div className="text-[13px] text-zinc-800 font-medium overflow-hidden text-ellipsis whitespace-nowrap">{entry.question}</div>
            <div className="text-[11px] text-zinc-400 mt-0.5 font-mono">
              {new Date(entry.timestamp).toLocaleString([], {
                month: "short",
                day: "numeric",
                hour: "numeric",
                minute: "2-digit",
              })}
            </div>
          </button>
        ))}
    </div>
  );
}
