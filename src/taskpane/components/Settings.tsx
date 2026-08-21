import { useState } from "react";
import { NoahHeader } from "./NoahShell";
import { SignOutModal } from "./SignOutModal";

export function Toggle({
  checked,
  onChange,
  id,
}: {
  checked: boolean;
  onChange: (v: boolean) => void;
  id?: string;
}) {
  return (
    <button
      id={id}
      type="button"
      role="switch"
      aria-checked={checked}
      onClick={() => onChange(!checked)}
      className={`relative inline-flex h-5 w-9 flex-shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none focus:ring-2 focus:ring-accent/30 ${
        checked ? "bg-accent" : "bg-border-strong/60 hover:bg-border-strong"
      }`}
    >
      <span
        className={`pointer-events-none inline-block h-4 w-4 transform rounded-full bg-white shadow-xs ring-0 transition duration-200 ease-in-out ${
          checked ? "translate-x-4" : "translate-x-0"
        }`}
      />
    </button>
  );
}

export function SettingsSection({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <div className="border-b border-border/50 pb-5 last:border-none last:pb-0">
      <h3 className="text-[11px] font-semibold text-ink-muted uppercase tracking-wider m-0 mb-3">
        {label}
      </h3>
      {children}
    </div>
  );
}

export function TextSizeControl({
  value,
  onChange,
  min = 12,
  max = 18,
}: {
  value: number;
  onChange: (v: number) => void;
  min?: number;
  max?: number;
}) {
  return (
    <div className="inline-flex items-center gap-1.5 p-1 bg-surface rounded-xl border border-border/70">
      <button
        type="button"
        className="w-7 h-7 rounded-lg border border-border/60 bg-canvas text-ink cursor-pointer inline-flex items-center justify-center text-[14px] font-medium hover:bg-surface-hover hover:border-border disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
        disabled={value <= min}
        onClick={() => onChange(Math.max(min, value - 1))}
        aria-label="Decrease text size"
      >
        −
      </button>
      <span className="text-[12.5px] font-medium text-ink min-w-[36px] text-center select-none">
        {value}px
      </span>
      <button
        type="button"
        className="w-7 h-7 rounded-lg border border-border/60 bg-canvas text-ink cursor-pointer inline-flex items-center justify-center text-[14px] font-medium hover:bg-surface-hover hover:border-border disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
        disabled={value >= max}
        onClick={() => onChange(Math.min(max, value + 1))}
        aria-label="Increase text size"
      >
        +
      </button>
    </div>
  );
}

function getInitials(emailStr: string): string {
  if (!emailStr) return "U";
  const namePart = emailStr.split("@")[0] || "";
  const parts = namePart.split(/[._\-\s]+/).filter(Boolean);
  if (parts.length >= 2) {
    return (parts[0][0] + parts[1][0]).toUpperCase();
  }
  if (namePart.length >= 2) {
    return namePart.substring(0, 2).toUpperCase();
  }
  return namePart.toUpperCase() || "U";
}

export function SettingsPage({
  onBack,
  email,
  useSelectionDefault,
  onToggleSelectionDefault,
  textSize: _textSize,
  onTextSizeChange: _onTextSizeChange,
  instructions: _instructions,
  onInstructionsChange: _onInstructionsChange,
  organizations,
  orgId,
  onOrgChange,
  onSignOut,
}: {
  onBack: () => void;
  email: string;
  useSelectionDefault: boolean;
  onToggleSelectionDefault: (v: boolean) => void;
  textSize?: number;
  onTextSizeChange?: (v: number) => void;
  instructions?: string;
  onInstructionsChange?: (v: string) => void;
  organizations: { id: string; name: string }[];
  orgId: string;
  onOrgChange: (id: string) => void;
  onSignOut: () => void;
}) {
  const [showConfirm, setShowConfirm] = useState(false);
  const initials = getInitials(email);

  return (
    <div className="flex flex-col flex-1 w-full h-full min-w-0 overflow-hidden bg-canvas">
      <NoahHeader title="Settings" onBack={onBack} />

      <div className="flex-auto overflow-y-auto overflow-x-hidden p-5 space-y-5">
        {/* Account */}
        <SettingsSection label="Account">
          <div className="flex items-center justify-between py-1">
            <div className="flex items-center gap-3 min-w-0 pr-3">
              <div className="w-8 h-8 rounded-full bg-accent-soft text-accent border border-accent/25 text-[12px] font-semibold flex items-center justify-center flex-none select-none">
                {initials}
              </div>
              <div className="flex flex-col gap-0.5 min-w-0">
                <span className="text-[13.5px] font-medium text-ink truncate">
                  {email || "User Account"}
                </span>
                <span className="text-[12px] text-ink-muted">Signed in</span>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setShowConfirm(true)}
              className="px-3 py-1.5 border border-red-600/30 rounded-lg bg-red-50 text-red-600 text-[12px] font-semibold cursor-pointer transition-all hover:bg-red-600 hover:text-white hover:border-red-600 flex-none shadow-2xs"
            >
              Sign out
            </button>
          </div>

          {organizations.length > 1 && (
            <div className="flex items-center justify-between pt-3 mt-3 border-t border-border/40">
              <span className="text-[13px] text-ink font-medium">Organization</span>
              <select
                className="h-8 border border-border/80 rounded-lg bg-canvas text-ink text-[12px] font-medium px-2.5 outline-none focus:border-accent cursor-pointer max-w-[170px] truncate"
                value={orgId}
                onChange={(e) => onOrgChange(e.target.value)}
              >
                {organizations.map((o) => (
                  <option key={o.id} value={o.id}>
                    {o.name}
                  </option>
                ))}
              </select>
            </div>
          )}
        </SettingsSection>

        {/* Context */}
        <SettingsSection label="Context">
          <div className="flex items-center justify-between py-1">
            <div className="flex flex-col gap-0.5 pr-4">
              <label
                htmlFor="toggle-selection"
                className="text-[13.5px] font-medium text-ink cursor-pointer"
              >
                Use selection by default
              </label>
              <span className="text-[12px] text-ink-muted leading-snug">
                Send selected document text automatically with questions
              </span>
            </div>
            <Toggle
              id="toggle-selection"
              checked={useSelectionDefault}
              onChange={onToggleSelectionDefault}
            />
          </div>
        </SettingsSection>
      </div>

      <SignOutModal
        isOpen={showConfirm}
        onClose={() => setShowConfirm(false)}
        onConfirm={onSignOut}
      />
    </div>
  );
}


