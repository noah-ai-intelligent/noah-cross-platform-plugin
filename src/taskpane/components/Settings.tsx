import { NoahHeader } from "./NoahShell";

export function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <span
      className={`w-8 h-[18px] rounded-full border-none relative flex-none cursor-pointer p-0 transition-colors ${checked ? "bg-accent" : "bg-border-strong"}`}
      role="switch"
      aria-checked={checked}
      onClick={() => onChange(!checked)}
    >
      <span className={`absolute top-[2px] left-[2px] w-[14px] h-[14px] rounded-full bg-canvas transition-transform duration-150 ease-out ${checked ? "translate-x-[14px]" : ""}`} />
    </span>
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
    <div className="py-4 border-b border-border last:border-none">
      <p className="text-[12px] font-semibold text-ink uppercase tracking-wide m-0 mb-2">{label}</p>
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
    <div className="flex items-center gap-3">
      <button
        className="w-7 h-7 rounded-full border border-border bg-canvas text-ink cursor-pointer inline-flex items-center justify-center text-[14px] leading-none hover:bg-surface-hover"
        disabled={value <= min}
        onClick={() => onChange(Math.max(min, value - 1))}
        aria-label="Decrease text size"
      >
        −
      </button>
      <span className="text-[12px] text-ink-muted min-w-[32px] text-center">{value}px</span>
      <button
        className="w-7 h-7 rounded-full border border-border bg-canvas text-ink cursor-pointer inline-flex items-center justify-center text-[14px] leading-none hover:bg-surface-hover"
        disabled={value >= max}
        onClick={() => onChange(Math.min(max, value + 1))}
        aria-label="Increase text size"
      >
        +
      </button>
    </div>
  );
}

export function SettingsPage({
  onBack,
  email,
  useSelectionDefault,
  onToggleSelectionDefault,
  textSize,
  onTextSizeChange,
  instructions,
  onInstructionsChange,
  organizations,
  orgId,
  onOrgChange,
  onSignOut,
}: {
  onBack: () => void;
  email: string;
  useSelectionDefault: boolean;
  onToggleSelectionDefault: (v: boolean) => void;
  textSize: number;
  onTextSizeChange: (v: number) => void;
  instructions: string;
  onInstructionsChange: (v: string) => void;
  organizations: { id: string; name: string }[];
  orgId: string;
  onOrgChange: (id: string) => void;
  onSignOut: () => void;
}) {
  return (
    <div className="flex flex-col flex-1 w-full overflow-hidden">
      <NoahHeader title="Settings" onBack={onBack} />
      <div className="flex-auto overflow-y-auto p-4">
        <SettingsSection label="Account">
          <div className="flex items-center justify-between py-2">
            <div className="flex flex-col gap-[2px]">
              <span className="text-[13px] text-ink">{email}</span>
              <span className="text-[12px] text-ink-muted">Signed in</span>
            </div>
            <button className="inline-flex items-center gap-1 h-[26px] px-2 border border-border rounded-full bg-canvas text-ink-secondary text-[12px] cursor-pointer hover:bg-surface-hover hover:text-ink" onClick={onSignOut}>
              Sign out
            </button>
          </div>
          {organizations.length > 1 && (
            <div className="flex items-center justify-between py-2 mt-3">
              <div className="flex flex-col gap-[2px]">
                <span className="text-[13px] text-ink">Organization</span>
              </div>
              <select
                className="h-[26px] border border-border rounded-full bg-canvas text-ink-secondary text-[12px] px-2 max-w-full"
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

        <SettingsSection label="Context">
          <div className="flex items-center justify-between py-2">
            <div className="flex flex-col gap-[2px]">
              <span className="text-[13px] text-ink">Use selection by default</span>
              <span className="text-[12px] text-ink-muted">
                Send what's selected in the document with each question
              </span>
            </div>
            <Toggle checked={useSelectionDefault} onChange={onToggleSelectionDefault} />
          </div>
        </SettingsSection>

        <SettingsSection label="Appearance">
          <div className="flex items-center justify-between py-2">
            <div className="flex flex-col gap-[2px]">
              <span className="text-[13px] text-ink">Text size</span>
            </div>
            <TextSizeControl value={textSize} onChange={onTextSizeChange} />
          </div>
        </SettingsSection>

        <SettingsSection label="Custom instructions">
          <textarea
            className="w-full min-h-[80px] border border-border rounded-[10px] p-2 px-3 text-[13px] text-ink bg-canvas resize-y outline-none focus:border-ink"
            placeholder="Tell NoahAI how you'd like it to respond…"
            value={instructions}
            onChange={(e) => onInstructionsChange(e.target.value)}
          />
          <p className="text-[12px] text-ink-muted mt-2">
            Saved on this device only.
          </p>
        </SettingsSection>
      </div>
    </div>
  );
}
