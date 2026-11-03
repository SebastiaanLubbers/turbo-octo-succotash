import type { ReactNode } from "react";
import { ArrowLeft } from "lucide-react";
import { Button } from "@/components/ui/button";

/** Shared layout for one wizard step: heading, body, and a footer with Back / primary action. */
export function StepFrame({
  title,
  children,
  onBack,
  backLabel = "Back",
  primary,
  secondary,
}: {
  title: string;
  children: ReactNode;
  onBack?: () => void;
  backLabel?: string;
  primary?: ReactNode;
  secondary?: ReactNode;
}) {
  return (
    <section aria-labelledby="step-title" className="flex min-h-full flex-col">
      <h1 id="step-title" tabIndex={-1} className="text-2xl font-semibold tracking-tight outline-none">
        {title}
      </h1>
      <div className="mt-6 flex-1 space-y-6">{children}</div>
      {(onBack || primary || secondary) && (
        <div className="mt-8 flex flex-wrap items-center gap-3 border-t pt-5">
          {onBack && (
            <Button variant="ghost" onClick={onBack}>
              <ArrowLeft aria-hidden="true" />
              {backLabel}
            </Button>
          )}
          <div className="ml-auto flex flex-wrap items-center gap-3">
            {secondary}
            {primary}
          </div>
        </div>
      )}
    </section>
  );
}

export function Field({ label, htmlFor, error, hint, children }: { label: string; htmlFor: string; error?: string; hint?: string; children: ReactNode }) {
  return (
    <div className="space-y-1.5">
      <label htmlFor={htmlFor} className="text-sm font-medium">
        {label}
      </label>
      {children}
      {hint && !error && <p id={`${htmlFor}-hint`} className="text-xs text-muted-foreground">{hint}</p>}
      {error && (
        <p id={`${htmlFor}-error`} className="text-sm text-red-700 dark:text-red-300">
          {error}
        </p>
      )}
    </div>
  );
}
