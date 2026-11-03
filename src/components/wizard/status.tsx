import type { FolderResultStatus, UploadStatus } from "@/models";
import { cn } from "@/lib/utils";

// One state-tone map for every status in the app.
type Tone = "success" | "info" | "danger" | "neutral" | "warning";

const TONE_CLASSES: Record<Tone, { badge: string; dot: string }> = {
  success: { badge: "bg-green-50 text-green-800 border-green-200 dark:bg-green-950 dark:text-green-200 dark:border-green-800", dot: "bg-green-700 dark:bg-green-400" },
  info: { badge: "bg-blue-50 text-blue-800 border-blue-200 dark:bg-blue-950 dark:text-blue-200 dark:border-blue-800", dot: "bg-blue-700 dark:bg-blue-400" },
  danger: { badge: "bg-red-50 text-red-800 border-red-200 dark:bg-red-950 dark:text-red-200 dark:border-red-800", dot: "bg-red-700 dark:bg-red-400" },
  warning: { badge: "bg-amber-50 text-amber-900 border-amber-200 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800", dot: "bg-amber-700 dark:bg-amber-400" },
  neutral: { badge: "bg-muted text-foreground border-border", dot: "bg-zinc-500" },
};

const STATUS_TONE: Record<FolderResultStatus | UploadStatus | "Pending" | "Working", Tone> = {
  Created: "success",
  "Already exists": "info",
  Failed: "danger",
  Skipped: "neutral",
  Uploaded: "success",
  Blocked: "warning",
  Pending: "neutral",
  Working: "info",
};

export function StatusBadge({ status, className }: { status: keyof typeof STATUS_TONE; className?: string }) {
  const tone = TONE_CLASSES[STATUS_TONE[status]];
  return (
    <span className={cn("inline-flex shrink-0 items-center gap-1.5 rounded-full border px-2 py-0.5 text-xs font-medium", tone.badge, className)}>
      <span className={cn("status-dot size-1.5 rounded-full", tone.dot)} aria-hidden="true" />
      {status}
    </span>
  );
}
