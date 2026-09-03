import type { ReactNode } from "react";
import logoUrl from "../../assets/icon-80.png";

export function NoahShell({ children }: { children: ReactNode }) {
  return <div className="flex flex-col h-screen w-full min-w-0 max-w-full overflow-hidden relative select-none">{children}</div>;
}

export function NoahHeader({
  title,
  subtitle,
  onBack,
  actions,
}: {
  title: string;
  subtitle?: ReactNode;
  onBack?: () => void;
  actions?: ReactNode;
}) {
  return (
    <div className="flex-none h-[52px] flex items-center justify-between px-4 border-b border-border">
      <div className="flex items-center gap-2 min-w-0">
        {onBack ? (
          <button className="w-[30px] h-[30px] inline-flex items-center justify-center rounded-lg border-none bg-transparent text-ink-secondary cursor-pointer hover:bg-surface-hover hover:text-ink" onClick={onBack} aria-label="Back">
            <BackIcon />
          </button>
        ) : (
          <img src={logoUrl} alt="Noah Logo" className="w-[22px] h-[22px] object-contain flex-none" onError={(e) => { e.currentTarget.style.display = 'none'; }} />
        )}
        <div className="flex flex-col min-w-0">
          <span className="font-semibold text-[14px] text-ink overflow-hidden text-ellipsis whitespace-nowrap">{title}</span>
          {subtitle && <div className="text-[11px] text-ink-secondary truncate">{subtitle}</div>}
        </div>
      </div>
      <div className="flex items-center gap-1 flex-none">{actions}</div>
    </div>
  );
}

export function NoahToolbar({ children }: { children: ReactNode }) {
  return <div className="flex-none min-h-[40px] flex items-center gap-2 py-2 px-4 border-b border-border flex-wrap">{children}</div>;
}

function BackIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
      <path
        d="M10 3L5 8l5 5"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}
