export function SignOutModal({
  isOpen,
  onClose,
  onConfirm,
}: {
  isOpen: boolean;
  onClose: () => void;
  onConfirm: () => void;
}) {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
      {/* Backdrop */}
      <div
        className="fixed inset-0 bg-black/30 transition-opacity animate-in fade-in duration-150"
        onClick={onClose}
        aria-hidden="true"
      />

      {/* Modal Content */}
      <div className="relative w-full max-w-[290px] bg-canvas border border-border rounded-2xl shadow-xl p-5 z-10 animate-in fade-in zoom-in-95 duration-150 text-center flex flex-col items-center box-border">
        <h3 className="text-[15px] font-semibold text-ink m-0 mb-1.5">
          Are you sure you want to sign out?
        </h3>

        <p className="text-[13px] text-ink-muted leading-normal m-0 mb-4">
          You’ll need to sign in again to access your Noah account.
        </p>

        <div className="flex items-center gap-2.5 w-full">
          <button
            type="button"
            className="flex-1 py-2 px-3 border border-border rounded-xl bg-canvas text-ink text-[13px] font-medium cursor-pointer transition-colors hover:bg-surface-hover active:bg-surface-hover/80"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="flex-1 py-2 px-3 border border-transparent rounded-xl bg-ink text-white text-[13px] font-semibold cursor-pointer transition-colors hover:bg-black active:bg-black shadow-xs"
            onClick={() => {
              onConfirm();
              onClose();
            }}
          >
            Sign out
          </button>
        </div>
      </div>
    </div>
  );
}
