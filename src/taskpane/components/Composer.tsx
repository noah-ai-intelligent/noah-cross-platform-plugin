import { SelectedContext } from "./Conversation";

export function SendButton({
  busy,
  disabled,
  onSend,
  onStop,
}: {
  busy: boolean;
  disabled: boolean;
  onSend: () => void;
  onStop: () => void;
}) {
  if (busy) {
    return (
      <button
        className="flex-none w-[32px] h-[32px] rounded-full border-none inline-flex items-center justify-center cursor-pointer bg-red-500 hover:bg-red-600 text-white transition-all shadow-xs active:scale-95 animate-pulse"
        onClick={onStop}
        aria-label="Stop generating"
        title="Stop response generation"
      >
        <StopIcon />
      </button>
    );
  }
  return (
    <button
      className={`flex-none w-[32px] h-[32px] rounded-full border-none inline-flex items-center justify-center cursor-pointer transition-all shadow-xs ${
        disabled
          ? "bg-zinc-100 text-zinc-300 cursor-not-allowed"
          : "bg-emerald-600 hover:bg-emerald-700 text-white active:scale-95 hover:shadow-md"
      }`}
      disabled={disabled}
      onClick={onSend}
      aria-label="Send message"
      title="Send message"
    >
      <SendIcon />
    </button>
  );
}

export function Composer({
  value,
  onChange,
  onSend,
  onStop,
  busy,
  showSelectionToggle,
  useSelection,
  onToggleSelection,
  selectionHint,
}: {
  value: string;
  onChange: (v: string) => void;
  onSend: () => void;
  onStop: () => void;
  busy: boolean;
  showSelectionToggle: boolean;
  useSelection: boolean;
  onToggleSelection: (v: boolean) => void;
  selectionHint: string;
}) {
  const canSend = value.trim().length > 0;

  return (
    <div className="flex-none bg-surface/80 backdrop-blur-xs border-t border-zinc-200/60 p-2.5">
      <div className="flex flex-col border border-zinc-200/90 rounded-2xl bg-white focus-within:border-emerald-500 focus-within:ring-2 focus-within:ring-emerald-500/20 shadow-xs focus-within:shadow-md transition-all overflow-hidden">
        {showSelectionToggle && (
          <SelectedContext enabled={useSelection} onToggle={onToggleSelection} hint={selectionHint} />
        )}
        <div className="flex items-end gap-2 p-2.5 min-h-[46px]">
          <textarea
            className="flex-1 border-none outline-none resize-none bg-transparent text-zinc-800 text-[13px] leading-relaxed max-h-[120px] py-1 px-1 placeholder:text-zinc-400"
            rows={2}
            placeholder="Ask NoahAI for help, or describe what you want to do…"
            value={value}
            onChange={(e) => onChange(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey && !busy) {
                e.preventDefault();
                if (canSend) onSend();
              }
            }}
          />
          <SendButton busy={busy} disabled={!canSend} onSend={onSend} onStop={onStop} />
        </div>
      </div>
    </div>
  );
}

function SendIcon() {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
      <line x1="12" y1="19" x2="12" y2="5"></line>
      <polyline points="5 12 12 5 19 12"></polyline>
    </svg>
  );
}

function StopIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
      <rect x="3.5" y="3.5" width="9" height="9" rx="1.5" fill="currentColor" />
    </svg>
  );
}

