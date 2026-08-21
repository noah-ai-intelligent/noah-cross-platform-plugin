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
  const btnClass = "flex-none w-[30px] h-[30px] rounded-full border-none inline-flex items-center justify-center cursor-pointer bg-ink text-canvas disabled:bg-disabled-bg disabled:text-disabled-ink disabled:cursor-not-allowed";

  if (busy) {
    return (
      <button className={btnClass} onClick={onStop} aria-label="Stop">
        <StopIcon />
      </button>
    );
  }
  return (
    <button className={btnClass} disabled={disabled} onClick={onSend} aria-label="Send">
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
    <div className="flex-none border-t border-border py-3 px-2">
      <div className="flex flex-col border border-border rounded-[14px] bg-canvas focus-within:border-ink transition-colors overflow-hidden">
        {showSelectionToggle && (
          <SelectedContext enabled={useSelection} onToggle={onToggleSelection} hint={selectionHint} />
        )}
        <div className="flex items-end gap-2 p-2 min-h-[44px]">
          <textarea
            className="flex-1 border-none outline-none resize-none bg-transparent text-ink text-[13px] leading-normal max-h-[120px] py-1 px-0.5"
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
    <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
      <path
        d="M8 13V3M8 3L3.5 7.5M8 3l4.5 4.5"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function StopIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
      <rect x="3" y="3" width="10" height="10" rx="2" fill="currentColor" />
    </svg>
  );
}
