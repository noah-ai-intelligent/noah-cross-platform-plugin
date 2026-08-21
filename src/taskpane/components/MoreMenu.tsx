import { useState } from "react";
import { SignOutModal } from "./SignOutModal";

export function MoreMenu({
  onClose,
  onSettings,
  onHistory,
  onSignOut,
}: {
  onClose: () => void;
  onSettings: () => void;
  onHistory: () => void;
  onSignOut: () => void;
}) {
  const [showConfirm, setShowConfirm] = useState(false);

  if (showConfirm) {
    return (
      <SignOutModal
        isOpen={true}
        onClose={() => {
          setShowConfirm(false);
          onClose();
        }}
        onConfirm={() => {
          onSignOut();
          onClose();
        }}
      />
    );
  }

  return (
    <>
      {/* Click outside backdrop */}
      <div
        className="fixed inset-0 z-40 bg-black/5"
        onClick={onClose}
        aria-hidden="true"
      />

      {/* Dropdown Menu */}
      <div
        className="absolute top-[48px] right-3 w-[190px] bg-canvas border border-border rounded-xl shadow-[0_4px_20px_-2px_rgba(0,0,0,0.12),0_2px_6px_-1px_rgba(0,0,0,0.06)] z-50 p-1.5 animate-in fade-in zoom-in-95 duration-100 focus:outline-none"
        role="menu"
      >
        <button
          role="menuitem"
          className="flex items-center gap-2.5 w-full text-left px-3 py-2 border-none rounded-lg bg-transparent text-ink text-[13px] font-medium cursor-pointer transition-colors duration-150 hover:bg-surface-hover active:bg-surface-hover/80 hover:text-ink focus:bg-surface-hover focus:outline-none"
          onClick={() => {
            onHistory();
            onClose();
          }}
        >
          <HistoryIcon className="w-4 h-4 text-ink-muted flex-none" />
          <span>History</span>
        </button>

        <button
          role="menuitem"
          className="flex items-center gap-2.5 w-full text-left px-3 py-2 border-none rounded-lg bg-transparent text-ink text-[13px] font-medium cursor-pointer transition-colors duration-150 hover:bg-surface-hover active:bg-surface-hover/80 hover:text-ink focus:bg-surface-hover focus:outline-none"
          onClick={() => {
            onSettings();
            onClose();
          }}
        >
          <SettingsIcon className="w-4 h-4 text-ink-muted flex-none" />
          <span>Settings</span>
        </button>

        <div className="my-1 border-t border-border/70" />

        <button
          role="menuitem"
          className="flex items-center gap-2.5 w-full text-left px-3 py-2 mt-1 border border-red-600/30 rounded-lg bg-red-50 text-red-600 text-[13px] font-semibold cursor-pointer transition-all duration-150 hover:bg-red-600 hover:text-white hover:border-red-600 group focus:outline-none shadow-2xs"
          onClick={() => {
            setShowConfirm(true);
          }}
        >
          <LogOutIcon className="w-4 h-4 text-red-600 group-hover:text-white flex-none transition-colors" />
          <span>Sign out</span>
        </button>
      </div>
    </>
  );
}

/* --- SVG Icons --- */

function HistoryIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
      <path d="M2.5 8a5.5 5.5 0 1 0 1.6-3.9L2.5 5.5" />
      <path d="M2.5 2.5v3h3" />
      <path d="M8 5v3.5l2.25 1.35" />
    </svg>
  );
}

function SettingsIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
      <path d="M2 4h4m3 0h5M2 8h8m3 0h1M2 12h2m3 0h7" />
      <circle cx="7.5" cy="4" r="1.5" />
      <circle cx="11.5" cy="8" r="1.5" />
      <circle cx="5.5" cy="12" r="1.5" />
    </svg>
  );
}

function LogOutIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
      <path d="M6 2.5H3.75A1.25 1.25 0 0 0 2.5 3.75v8.5c0 .69.56 1.25 1.25 1.25H6" />
      <path d="M10.25 11.25L13.5 8l-3.25-3.25" />
      <path d="M13.5 8H6" />
    </svg>
  );
}

