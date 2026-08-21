import type { AddonAnswer, Citation } from "../../addonClient";
import { describeAnchor } from "../../document/anchor";
import ReactMarkdown from 'react-markdown';

export type HistoryEntry = {
  id: string;
  question: string;
  preview: string;
  timestamp: number;
};

export function CellReference({
  citation,
  onClick,
}: {
  citation: Citation;
  onClick: () => void;
}) {
  return (
    <button className="inline-flex items-center gap-1 border border-border rounded-md bg-surface text-ink-secondary text-[12px] py-[2px] px-2 cursor-pointer font-mono hover:bg-surface-hover hover:text-ink" onClick={onClick} title="Show me where this came from">
      {describeAnchor(citation.anchor)}
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
    <div className="flex items-center justify-between bg-surface-hover p-2 px-3 text-[12px] text-ink-secondary border-b border-border">
      <div className="flex items-center gap-2">
        <span>{hint}</span>
        <span>Selected</span>
      </div>

      <button className="inline-flex items-center justify-center border-none bg-transparent text-ink-muted cursor-pointer p-1 hover:text-ink" onClick={() => onToggle(false)} aria-label="Remove selection">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
          <line x1="18" y1="6" x2="6" y2="18"></line>
          <line x1="6" y1="6" x2="18" y2="18"></line>
        </svg>
      </button>
    </div>
  );
}

export function ActionBar({ children }: { children: React.ReactNode }) {
  return <div className="flex gap-1 flex-wrap mt-1">{children}</div>;
}

export function LoadingState({ activity }: { activity: string }) {
  return (
    <div className="flex items-center gap-2 text-ink-muted text-[13px] p-1">
      <span className="inline-flex gap-[3px] animate-pulse">
        <span className="w-[5px] h-[5px] rounded-full bg-ink-muted" />
        <span className="w-[5px] h-[5px] rounded-full bg-ink-muted" />
        <span className="w-[5px] h-[5px] rounded-full bg-ink-muted" />
      </span>
      {activity || "Thinking…"}
    </div>
  );
}

export function ErrorState({ message }: { message: string }) {
  return <div className="flex items-start gap-2 border border-[#fecaca] bg-danger-soft text-danger rounded-[10px] p-2 px-3 text-[13px]">{message}</div>;
}

export function UserMessage({ text }: { text: string }) {
  return (
    <div className="flex flex-col gap-1 items-end">
      <div className="max-w-[90%] text-[13px] whitespace-pre-wrap break-words bg-surface border border-border rounded-[14px] py-2 px-3">
        <span style={{ whiteSpace: "pre-wrap" }}>{text}</span>
      </div>
    </div>
  );
}

import type { ChatMessage } from "../../chatClient";

export function AssistantMessage({
  message,
  answer,
  host,
  onInsertProse,
  onCitation,
}: {
  message: ChatMessage;
  answer?: AddonAnswer | null;
  host: string;
  onInsertProse: (text: string) => void;
  onCitation?: (c: Citation) => void;
}) {
  const textBlocks = message.content.filter((b) => b.type === "text" && typeof b.text === "string");
  const fullText = textBlocks.map((b) => b.text).join("\n\n");

  return (
    <div className="flex flex-col gap-1">
      <div className="max-w-[90%] text-[13px] whitespace-pre-wrap break-words text-ink px-1">
        {answer && answer.blocks && answer.blocks.length > 0 ? (
          answer.blocks.map((block, i) => (
            <div key={i} style={{ marginBottom: 6 }}>
              <ReactMarkdown>{block.text}</ReactMarkdown>
              {block.citations && block.citations.map((index) => {
                const citation = answer.citations[index];
                if (!citation) return null;
                return <CellReference key={index} citation={citation} onClick={() => onCitation && onCitation(citation)} />;
              })}
            </div>
          ))
        ) : (
          <div style={{ marginBottom: 6 }}>
            <ReactMarkdown>{answer?.markdown || answer?.prose || fullText}</ReactMarkdown>
          </div>
        )}
        {host !== "Excel" && ((answer && answer.prose) || fullText) && (
          <ActionBar>
            <button className="inline-flex items-center gap-1 h-[26px] px-2 border border-border rounded-full bg-canvas text-ink-secondary text-[12px] cursor-pointer hover:bg-surface-hover hover:text-ink" onClick={() => onInsertProse(answer ? answer.prose : fullText)}>
              Insert into document
            </button>
          </ActionBar>
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
  onCitation,
}: {
  messages: ChatMessage[];
  latestAnswer?: AddonAnswer | null;
  busy: boolean;
  activity: string;
  error: string | null;
  host: string;
  onInsertProse: (text: string) => void;
  onCitation?: (c: Citation) => void;
}) {
  return (
    <div className="flex flex-col gap-4 p-4">
      {messages.map((msg, index) => {
        const isLast = index === messages.length - 1;
        return msg.role === "user" ? (
          <UserMessage key={msg.id} text={msg.content[0]?.text || ""} />
        ) : (
          <AssistantMessage
            key={msg.id}
            message={msg}
            answer={isLast && latestAnswer && !busy ? latestAnswer : null}
            host={host}
            onInsertProse={onInsertProse}
            onCitation={onCitation}
          />
        );
      })}

      {busy && latestAnswer && (
        <AssistantMessage
          message={{ id: "streaming", role: "assistant", content: [] }}
          answer={latestAnswer}
          host={host}
          onInsertProse={onInsertProse}
          onCitation={onCitation}
        />
      )}

      {busy && !latestAnswer && <LoadingState activity={activity} />}
      {!busy && error && <ErrorState message={error} />}
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
      <div className="flex flex-col items-center justify-center gap-2 h-full text-ink-muted text-[13px] text-center py-8 px-5">
        <span>No conversations yet</span>
        <span className="text-[12px] text-ink-muted mt-[2px]">Questions you ask NoahAI will show up here.</span>
      </div>
    );
  }
  return (
    <div>
      {entries
        .slice()
        .reverse()
        .map((entry) => (
          <button key={entry.id} className="block w-full text-left border-none border-b border-border bg-transparent py-3 px-0 cursor-pointer overflow-hidden hover:bg-surface-hover" onClick={() => onSelect(entry)}>
            <div className="text-[13px] text-ink font-medium overflow-hidden text-ellipsis whitespace-nowrap">{entry.question}</div>
            <div className="text-[12px] text-ink-muted mt-[2px]">
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
