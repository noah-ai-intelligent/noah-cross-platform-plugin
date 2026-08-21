import { useEffect, useState, useMemo } from "react";
import { NoahHeader } from "./NoahShell";
import { listConversations, type Conversation } from "../../chatClient";

interface HistoryPageProps {
  orgId: string;
  currentId?: string;
  onBack: () => void;
  onSelect: (conv: Conversation) => void;
  onNewChat?: () => void;
}

interface GroupedConversations {
  label: string;
  items: Conversation[];
}

export function HistoryPage({
  orgId,
  currentId,
  onBack,
  onSelect,
  onNewChat,
}: HistoryPageProps) {
  const [entries, setEntries] = useState<Conversation[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState("");

  useEffect(() => {
    if (!orgId) return;
    setLoading(true);
    listConversations(orgId, 50)
      .then((data) => setEntries(data || []))
      .catch((e) => console.error("Failed to load history", e))
      .finally(() => setLoading(false));
  }, [orgId]);

  // Filter entries based on search query
  const filteredEntries = useMemo(() => {
    if (!searchQuery.trim()) return entries;
    const q = searchQuery.toLowerCase().trim();
    return entries.filter((e) =>
      (e.title || "New Conversation").toLowerCase().includes(q)
    );
  }, [entries, searchQuery]);

  // Group entries by date
  const groupedEntries = useMemo(() => {
    return groupConversationsByDate(filteredEntries);
  }, [filteredEntries]);

  return (
    <div className="flex flex-col flex-1 w-full h-full min-w-0 overflow-hidden bg-canvas">
      <NoahHeader
        title="History"
        onBack={onBack}
        actions={
          onNewChat ? (
            <button
              onClick={onNewChat}
              className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[12px] font-medium text-accent bg-accent-soft hover:bg-accent hover:text-white rounded-lg border border-accent/30 transition-all cursor-pointer shadow-2xs"
              title="Start a new chat"
            >
              <PlusIcon className="w-3.5 h-3.5" />
              <span>New Chat</span>
            </button>
          ) : undefined
        }
      />

      {/* Search Bar & Subheader Stats */}
      <div className="flex-none px-4 py-3 border-b border-border/60 bg-canvas space-y-2">
        <div className="relative flex items-center w-full">
          <SearchIcon className="absolute left-3 w-4 h-4 text-ink-muted pointer-events-none flex-none" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search conversations..."
            className="w-full pl-9 pr-8 py-2 bg-surface hover:bg-surface-hover/80 focus:bg-canvas border border-border focus:border-accent rounded-xl text-[13px] text-ink placeholder:text-ink-placeholder transition-all outline-none focus:ring-2 focus:ring-accent/20 box-border"
          />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery("")}
              className="absolute right-2.5 p-1 rounded-md text-ink-muted hover:text-ink hover:bg-surface-hover border-none bg-transparent cursor-pointer flex-none transition-colors"
              title="Clear search"
            >
              <CloseIcon className="w-3.5 h-3.5" />
            </button>
          )}
        </div>

        {/* Stats summary bar */}
        {!loading && entries.length > 0 && (
          <div className="flex items-center justify-between text-[11.5px] text-ink-muted px-0.5">
            <span>
              {searchQuery ? (
                <>Found <strong className="text-ink font-semibold">{filteredEntries.length}</strong> matching conversation{filteredEntries.length !== 1 ? "s" : ""}</>
              ) : (
                <><strong className="text-ink font-semibold">{entries.length}</strong> total conversation{entries.length !== 1 ? "s" : ""}</>
              )}
            </span>
            {searchQuery && (
              <button
                onClick={() => setSearchQuery("")}
                className="text-accent hover:underline cursor-pointer border-none bg-transparent p-0 text-[11.5px] font-medium"
              >
                Reset search
              </button>
            )}
          </div>
        )}
      </div>

      {/* Content List Area */}
      <div className="flex-auto overflow-y-auto overflow-x-hidden px-4 py-3">
        {loading ? (
          <HistorySkeletonLoader />
        ) : !entries.length ? (
          /* Empty State */
          <div className="flex flex-col items-center justify-center text-center py-16 px-4 h-full">
            <div className="w-12 h-12 rounded-2xl bg-accent-soft border border-accent/20 flex items-center justify-center text-accent mb-3.5 shadow-sm">
              <ChatBubbleIcon className="w-6 h-6" />
            </div>
            <span className="text-[15px] font-semibold text-ink mb-1">No conversations yet</span>
            <p className="text-[12.5px] text-ink-muted max-w-[240px] leading-relaxed m-0">
              Questions and context analysis you ask Noah will be saved here.
            </p>
            {onNewChat && (
              <button
                onClick={onNewChat}
                className="mt-5 px-4 py-2 bg-accent hover:bg-accent-hover text-white text-[13px] font-medium rounded-xl border-none cursor-pointer transition-all shadow-xs inline-flex items-center gap-1.5"
              >
                <PlusIcon className="w-4 h-4" />
                <span>Start New Chat</span>
              </button>
            )}
          </div>
        ) : !filteredEntries.length ? (
          /* Search Empty State */
          <div className="flex flex-col items-center justify-center text-center py-12 px-4">
            <div className="w-10 h-10 rounded-xl bg-surface border border-border/80 flex items-center justify-center text-ink-muted mb-3">
              <SearchIcon className="w-5 h-5" />
            </div>
            <span className="text-[14px] font-semibold text-ink mb-1">No matching conversations</span>
            <span className="text-[12px] text-ink-muted mb-4 max-w-[220px]">
              No results found for &ldquo;<span className="text-ink font-medium">{searchQuery}</span>&rdquo;
            </span>
            <button
              onClick={() => setSearchQuery("")}
              className="px-3.5 py-1.5 bg-surface hover:bg-surface-hover text-ink text-[12.5px] font-medium rounded-lg border border-border cursor-pointer transition-colors"
            >
              Clear Search
            </button>
          </div>
        ) : (
          /* Date-Grouped History List */
          <div className="space-y-5 w-full pb-4">
            {groupedEntries.map((group) => (
              <div key={group.label} className="space-y-2">
                {/* Group Title Header */}
                <div className="flex items-center gap-2 px-1 select-none">
                  <span className="text-[11px] font-semibold text-ink-muted uppercase tracking-wider">
                    {group.label}
                  </span>
                  <div className="h-px flex-1 bg-border/40" />
                  <span className="text-[10.5px] text-ink-muted/80 font-medium px-1.5 py-0.2 bg-surface rounded-full border border-border/40">
                    {group.items.length}
                  </span>
                </div>

                {/* Group Items */}
                <div className="space-y-1.5 w-full">
                  {group.items.map((entry) => {
                    const isSelected = entry.id === currentId;
                    return (
                      <button
                        key={entry.id}
                        onClick={() => onSelect(entry)}
                        className={`group relative flex items-center justify-between w-full p-3 rounded-xl border text-left cursor-pointer transition-all duration-150 min-w-0 box-border ${
                          isSelected
                            ? "bg-accent-soft/60 border-accent/40 border-l-4 border-l-accent shadow-2xs"
                            : "bg-canvas border-border/70 hover:bg-accent-soft/25 hover:border-accent/40 hover:border-l-4 hover:border-l-accent/70 hover:shadow-xs hover:-translate-y-[1px]"
                        }`}
                      >
                        {/* Icon + Title + Subtitle */}
                        <div className="flex items-center gap-3 min-w-0 flex-1 pr-2">
                          <div
                            className={`w-8 h-8 rounded-lg flex items-center justify-center flex-none transition-all duration-150 ${
                              isSelected
                                ? "bg-accent text-white shadow-2xs"
                                : "bg-surface text-ink-muted group-hover:bg-accent group-hover:text-white group-hover:shadow-2xs group-hover:scale-105 border border-border/50"
                            }`}
                          >
                            <ChatBubbleIcon className="w-4 h-4" />
                          </div>

                          <div className="flex flex-col min-w-0 flex-1 gap-0.5">
                            <div className="flex items-center gap-1.5 w-full min-w-0">
                              <span
                                className={`text-[13px] leading-tight truncate flex-1 min-w-0 transition-colors ${
                                  isSelected
                                    ? "font-semibold text-accent"
                                    : "font-medium text-ink group-hover:text-accent group-hover:font-semibold"
                                }`}
                                title={entry.title || "New Conversation"}
                              >
                                {entry.title || "New Conversation"}
                              </span>
                            </div>
                            <span className="text-[11px] text-ink-muted font-normal truncate group-hover:text-ink-secondary transition-colors">
                              {formatConversationTime(entry.create_time)}
                            </span>
                          </div>
                        </div>

                        {/* Right Action / Status Badge */}
                        <div className="flex items-center gap-1.5 flex-none">
                          {isSelected ? (
                            <span className="text-[9.5px] bg-accent text-white font-semibold px-2 py-0.5 rounded-full uppercase tracking-wider shadow-2xs">
                              Active
                            </span>
                          ) : (
                            <ChevronRightIcon className="w-4 h-4 text-ink-muted/40 group-hover:text-accent group-hover:translate-x-1 transition-all opacity-70 group-hover:opacity-100" />
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// Helpers
function groupConversationsByDate(items: Conversation[]): GroupedConversations[] {
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const yesterdayStart = todayStart - 86400000;
  const sevenDaysAgo = todayStart - 6 * 86400000;
  const thirtyDaysAgo = todayStart - 29 * 86400000;

  const today: Conversation[] = [];
  const yesterday: Conversation[] = [];
  const last7Days: Conversation[] = [];
  const last30Days: Conversation[] = [];
  const older: Conversation[] = [];

  items.forEach((item) => {
    const time = new Date(item.create_time).getTime();
    if (isNaN(time)) {
      older.push(item);
    } else if (time >= todayStart) {
      today.push(item);
    } else if (time >= yesterdayStart) {
      yesterday.push(item);
    } else if (time >= sevenDaysAgo) {
      last7Days.push(item);
    } else if (time >= thirtyDaysAgo) {
      last30Days.push(item);
    } else {
      older.push(item);
    }
  });

  const groups: GroupedConversations[] = [];
  if (today.length) groups.push({ label: "Today", items: today });
  if (yesterday.length) groups.push({ label: "Yesterday", items: yesterday });
  if (last7Days.length) groups.push({ label: "Previous 7 Days", items: last7Days });
  if (last30Days.length) groups.push({ label: "Previous 30 Days", items: last30Days });
  if (older.length) groups.push({ label: "Older", items: older });

  return groups;
}

function formatConversationTime(dateStr: string): string {
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    const now = new Date();

    const isToday =
      date.getDate() === now.getDate() &&
      date.getMonth() === now.getMonth() &&
      date.getFullYear() === now.getFullYear();

    if (isToday) {
      return date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    }

    const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
    const isYesterday =
      date.getDate() === yesterday.getDate() &&
      date.getMonth() === yesterday.getMonth() &&
      date.getFullYear() === yesterday.getFullYear();

    if (isYesterday) {
      return `Yesterday at ${date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
    }

    return date.toLocaleDateString([], { month: "short", day: "numeric" });
  } catch {
    return dateStr;
  }
}

// Grouped Skeleton Loader
function HistorySkeletonLoader() {
  return (
    <div className="space-y-5 w-full">
      {/* Group 1 Skeleton */}
      <div className="space-y-2">
        <div className="w-16 h-3 bg-surface-hover/80 rounded animate-pulse" />
        <div className="space-y-1.5">
          {[1, 2].map((i) => (
            <div
              key={i}
              className="flex items-center gap-3 p-3 rounded-xl border border-border/50 bg-surface/30 animate-pulse w-full"
            >
              <div className="w-8 h-8 rounded-lg bg-surface-hover flex-none" />
              <div className="flex-1 space-y-1.5">
                <div className="w-2/3 h-3 bg-surface-hover rounded" />
                <div className="w-1/3 h-2.5 bg-surface-hover/70 rounded" />
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Group 2 Skeleton */}
      <div className="space-y-2">
        <div className="w-24 h-3 bg-surface-hover/80 rounded animate-pulse" />
        <div className="space-y-1.5">
          {[3, 4, 5].map((i) => (
            <div
              key={i}
              className="flex items-center gap-3 p-3 rounded-xl border border-border/50 bg-surface/30 animate-pulse w-full"
            >
              <div className="w-8 h-8 rounded-lg bg-surface-hover flex-none" />
              <div className="flex-1 space-y-1.5">
                <div className="w-3/4 h-3 bg-surface-hover rounded" />
                <div className="w-1/4 h-2.5 bg-surface-hover/70 rounded" />
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// SVG Icons
function SearchIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor">
      <path
        d="M7.333 12.667A5.333 5.333 0 1 0 7.333 2a5.333 5.333 0 0 0 0 10.667zM14 14l-2.9-2.9"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function CloseIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor">
      <path d="M12 4L4 12M4 4l8 8" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function PlusIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor">
      <path d="M8 3.5v9M3.5 8h9" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function ChatBubbleIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor">
      <path
        d="M2.5 7.667c0-2.394 2.462-4.334 5.5-4.334s5.5 1.94 5.5 4.334c0 2.393-2.462 4.333-5.5 4.333-.62 0-1.218-.08-1.776-.23L3.5 12.5l.654-1.96C3.12 9.71 2.5 8.74 2.5 7.667z"
        strokeWidth="1.3"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function ChevronRightIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor">
      <path d="M6 3.5L10.5 8L6 12.5" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}









