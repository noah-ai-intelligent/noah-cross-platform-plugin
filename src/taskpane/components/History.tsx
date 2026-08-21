import { useEffect, useState } from "react";
import { NoahHeader } from "./NoahShell";
import { listConversations, type Conversation } from "../../chatClient";

export function HistoryPage({
  orgId,
  currentId,
  onBack,
  onNewChat,
  onSelect,
}: {
  orgId: string;
  currentId?: string;
  onBack: () => void;
  onNewChat?: () => void;
  onSelect: (conv: Conversation) => void;
}) {
  const [entries, setEntries] = useState<Conversation[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!orgId) return;
    setLoading(true);
    listConversations(orgId, 50)
      .then((data) => setEntries(data || []))
      .catch((e) => console.error("Failed to load history", e))
      .finally(() => setLoading(false));
  }, [orgId]);

  return (
    <div className="flex flex-col flex-1 w-full overflow-hidden">
      <NoahHeader 
        title="History" 
        onBack={onBack} 
        actions={
          onNewChat && (
            <button className="w-[30px] h-[30px] inline-flex items-center justify-center rounded-lg border-none bg-transparent text-ink-secondary cursor-pointer hover:bg-surface-hover hover:text-ink" onClick={onNewChat} aria-label="New Chat">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M8 3v10M3 8h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          )
        }
      />
      <div className="flex-auto overflow-y-auto p-4">
        {loading ? (
          <div className="flex items-center gap-2 text-ink-muted text-[13px] p-1">Loading...</div>
        ) : !entries.length ? (
          <div className="flex flex-col items-center justify-center gap-2 h-full text-ink-muted text-[13px] text-center py-8 px-5">
            <span>No conversations yet</span>
            <span className="text-[12px] text-ink-muted mt-[2px]">Questions you ask NoahAI will show up here.</span>
          </div>
        ) : (
          <div>
            {entries.map((entry) => (
              <button key={entry.id} className={`block w-full text-left border-none border-b border-border py-3 px-3 cursor-pointer overflow-hidden hover:bg-surface-hover ${entry.id === currentId ? 'bg-surface-hover' : 'bg-transparent'}`} onClick={() => onSelect(entry)}>
                <div className="text-[13px] text-ink font-medium overflow-hidden text-ellipsis whitespace-nowrap">{entry.title || "New Conversation"}</div>
                <div className="text-[12px] text-ink-muted mt-[2px]">
                  {new Date(entry.create_time).toLocaleString([], {
                    month: "short",
                    day: "numeric",
                    hour: "numeric",
                    minute: "2-digit",
                  })}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
