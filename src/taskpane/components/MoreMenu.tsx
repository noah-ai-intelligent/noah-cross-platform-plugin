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
  return (
    <>
      <div className="fixed inset-0 z-10" onClick={onClose} />
      <div className="absolute top-[56px] right-3 bg-canvas border border-border rounded-[10px] min-w-[180px] z-[11] overflow-hidden">
        <button
          className="flex items-center gap-2 w-full text-left p-2 px-3 border-none bg-transparent text-ink text-[13px] cursor-pointer hover:bg-surface-hover"
          onClick={() => {
            onHistory();
            onClose();
          }}
        >
          History
        </button>
        <button
          className="flex items-center gap-2 w-full text-left p-2 px-3 border-none bg-transparent text-ink text-[13px] cursor-pointer hover:bg-surface-hover"
          onClick={() => {
            onSettings();
            onClose();
          }}
        >
          Settings
        </button>
        <button
          className="flex items-center gap-2 w-full text-left p-2 px-3 border-none bg-transparent text-danger text-[13px] cursor-pointer hover:bg-surface-hover"
          onClick={() => {
            onSignOut();
            onClose();
          }}
        >
          Sign out
        </button>
      </div>
    </>
  );
}
