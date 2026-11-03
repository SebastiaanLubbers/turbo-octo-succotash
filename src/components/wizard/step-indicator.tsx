import { Check } from "lucide-react";
import { STEPS, type StepId, useWizard } from "@/state/wizard-store";
import { cn } from "@/lib/utils";

const GROUPS = ["Setup", "Structure", "Document"] as const;

/** Visible wizard progress. Completed steps can be revisited; later steps cannot be skipped to. */
export function StepIndicator({ canVisit, onVisit }: { canVisit: (id: StepId) => boolean; onVisit: (id: StepId) => void }) {
  const step = useWizard((s) => s.step);
  const current = STEPS.find((s) => s.id === step);
  return (
    <nav aria-label="Wizard progress" className="space-y-5">
      <p className="text-sm font-medium" aria-live="polite">
        Step {step} of {STEPS.length}
        <span className="sr-only">: {current?.label}</span>
      </p>
      <div className="h-1.5 overflow-hidden rounded-full bg-muted" aria-hidden="true">
        <div className="h-full rounded-full bg-primary transition-[width]" style={{ width: `${(step / STEPS.length) * 100}%` }} />
      </div>
      {GROUPS.map((group) => (
        <div key={group}>
          <p className="mb-1.5 text-xs font-semibold tracking-wide text-muted-foreground uppercase">{group}</p>
          <ol className="space-y-0.5">
            {STEPS.filter((s) => s.group === group).map((s) => {
              const state = s.id === step ? "current" : s.id < step ? "done" : "todo";
              const enabled = s.id !== step && canVisit(s.id);
              return (
                <li key={s.id}>
                  <button
                    type="button"
                    disabled={!enabled}
                    onClick={() => onVisit(s.id)}
                    aria-current={state === "current" ? "step" : undefined}
                    data-step-state={state}
                    className={cn(
                      "flex w-full items-center gap-2.5 rounded-md px-2 py-1.5 text-left text-sm transition-colors focus-visible:outline-2 focus-visible:outline-ring",
                      state === "current" && "bg-accent font-medium",
                      enabled && "hover:bg-accent",
                      !enabled && state !== "current" && "cursor-default",
                    )}
                  >
                    <span
                      className={cn(
                        "flex size-5 shrink-0 items-center justify-center rounded-full border text-[11px] tabular-nums",
                        state === "current" && "border-primary bg-primary text-primary-foreground",
                        state === "done" && "border-primary text-primary",
                        state === "todo" && "border-border text-muted-foreground",
                      )}
                      aria-hidden="true"
                    >
                      {state === "done" ? <Check className="size-3" /> : s.id}
                    </span>
                    <span className={cn(state === "todo" && "text-muted-foreground")}>{s.label}</span>
                    {state === "done" && <span className="sr-only">(completed)</span>}
                  </button>
                </li>
              );
            })}
          </ol>
        </div>
      ))}
    </nav>
  );
}
