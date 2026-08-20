type Skill = { command: string; label: string };

const SKILLS: Skill[] = [
  { command: "/analyze-sheet", label: "Analyze this sheet" },
  { command: "/clean-data", label: "Clean up messy data" },
  { command: "/formula-help", label: "Help me write a formula" },
  { command: "/summarize-data", label: "Summarize what's here" },
  { command: "/create-chart", label: "Create a chart" },
  { command: "/find-errors", label: "Find errors" },
  { command: "/format-sheet", label: "Format this sheet" },
];

export function WelcomeScreen({
  subtitle,
  children,
}: {
  subtitle: string;
  children?: React.ReactNode;
}) {
  return (
    <div className="h-full flex flex-col items-center justify-center text-center py-8 px-5 gap-2">
      <div className="w-[44px] h-[44px] rounded-[10px] bg-ink text-canvas flex items-center justify-center font-bold text-[16px] mb-2">N</div>
      <h2 className="text-[19px] font-bold text-ink m-0">Noah, right in your workbook</h2>
      <p className="text-[13px] text-ink-muted m-0 max-w-[280px]">{subtitle}</p>
      {children}
    </div>
  );
}

export function SkillSuggestions({ onPick }: { onPick: (command: string) => void }) {
  return (
    <div className="flex flex-col gap-2 w-full mt-6">
      {SKILLS.map((skill) => (
        <button
          key={skill.command}
          className="flex items-center gap-2 w-full text-left py-2 px-3 border border-border rounded-full bg-canvas text-ink text-[13px] cursor-pointer hover:bg-surface-hover hover:border-border-strong"
          onClick={() => onPick(skill.command + " ")}
        >
          <span>{skill.label}</span>
          <span className="text-ink-muted text-[12px]">{skill.command}</span>
        </button>
      ))}
    </div>
  );
}
