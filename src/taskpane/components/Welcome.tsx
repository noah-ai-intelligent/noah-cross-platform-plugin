import type React from "react";

export function WelcomeScreen({
  subtitle,
  children,
}: {
  subtitle: string;
  children?: React.ReactNode;
}) {
  const isSigned = subtitle.startsWith("Signed in as ");
  let email = "";

  if (isSigned) {
    const raw = subtitle.replace("Signed in as ", "");
    const parts = raw.split(" · ");
    email = parts[0] || "";
  }

  const displayEmail = email || (isSigned ? "" : subtitle);

  return (
    <div className="h-full flex flex-col items-center justify-center text-center py-10 px-6 animate-fadeIn">
      <h2 className="text-[20px] font-bold text-ink tracking-tight m-0 mb-2">
        Noah, right in your workbook
      </h2>

      {displayEmail ? (
        <div className="inline-flex items-center px-3 py-1 rounded-full bg-surface border border-border text-[12px] text-ink-secondary font-medium mb-3 shadow-2xs">
          {displayEmail}
        </div>
      ) : null}

      <p className="text-[13px] text-ink-muted m-0 max-w-[280px]">
        Ask questions, analyze data, and generate insights instantly.
      </p>

      {children}
    </div>
  );
}

export function SkillSuggestions(_props: { onPick: (command: string) => void }) {
  return null;
}
