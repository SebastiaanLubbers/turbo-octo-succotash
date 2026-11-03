import { useState } from "react";
import { AlertTriangle, ChevronDown } from "lucide-react";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";
import { AppError } from "@/services/errors";
import { cn } from "@/lib/utils";

/** Business message first; developer diagnostic separated under "Details for support". */
export function ErrorPanel({ error, title = "Something needs attention", action }: { error: unknown; title?: string; action?: React.ReactNode }) {
  const [open, setOpen] = useState(false);
  const appError = error instanceof AppError ? error : null;
  const message = appError?.userMessage ?? "Something went wrong. Try again.";
  return (
    <div role="alert" className="rounded-lg border border-red-200 bg-red-50 p-4 text-red-900 dark:border-red-800 dark:bg-red-950 dark:text-red-100">
      <div className="flex gap-3">
        <AlertTriangle className="mt-0.5 size-5 shrink-0" aria-hidden="true" />
        <div className="min-w-0 flex-1 space-y-2">
          <p className="font-medium">{title}</p>
          <p className="text-sm">{message}</p>
          {action}
          {appError?.diagnostic && (
            <Collapsible open={open} onOpenChange={setOpen}>
              <CollapsibleTrigger className="inline-flex items-center gap-1 rounded text-xs font-medium underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring">
                Details for support
                <ChevronDown className={cn("size-3.5 transition-transform", open && "rotate-180")} aria-hidden="true" />
              </CollapsibleTrigger>
              <CollapsibleContent>
                <code className="mt-1 block rounded bg-white/70 px-2 py-1 font-mono text-xs break-all text-red-950 dark:bg-black/30 dark:text-red-100">{appError.diagnostic}</code>
              </CollapsibleContent>
            </Collapsible>
          )}
        </div>
      </div>
    </div>
  );
}
