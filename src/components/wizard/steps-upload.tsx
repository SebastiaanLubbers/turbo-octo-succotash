import { useRef, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import {
  AlertTriangle,
  ArrowRight,
  CheckCircle2,
  ExternalLink,
  FileUp,
  FolderPlus,
  Upload,
  XCircle,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Spinner } from "@/components/ui/spinner";
import type { FolderResult } from "@/models";
import {
  SharePointFileService,
  fileServerPath,
  validateLocalFile,
  type UploadStage,
} from "@/services/sharepoint-file-service";
import { toAppError } from "@/services/errors";
import { formatBytes } from "@/services/validation";
import { sharePointLink, splitRel } from "@/state/selectors";
import { useWizard } from "@/state/wizard-store";
import { queryClient } from "@/lib/query-client";
import { safeUrl } from "@/lib/safe-url";
import { ErrorPanel } from "./error-panel";
import { PathLine } from "./destination-tree";
import { StepFrame } from "./step-frame";
import { StatusBadge } from "./status";

function useUploadTargets(): FolderResult[] {
  const results = useWizard((s) => s.folderResults);
  return ((results ?? []).filter(Boolean) as FolderResult[]).filter(
    (r) =>
      !r.isRoot && (r.status === "Created" || r.status === "Already exists"),
  );
}

function fileTypeLabel(file: File): string {
  const ext = file.name.includes(".")
    ? file.name.split(".").pop()!.toUpperCase()
    : "";
  return (
    [ext && `${ext} file`, file.type].filter(Boolean).join(" · ") ||
    "Unknown type"
  );
}

// ---------- 10. Upload document ----------
export function UploadStep() {
  const file = useWizard((s) => s.file);
  const setFile = useWizard((s) => s.setFile);
  const destination = useWizard((s) => s.destinationServerPath);
  const setDestination = useWizard((s) => s.setDestination);
  const goTo = useWizard((s) => s.goTo);
  const targets = useUploadTargets();
  const inputRef = useRef<HTMLInputElement>(null);
  const [submitted, setSubmitted] = useState(false);
  const fileProblem = validateLocalFile(file);
  const destinationValid = targets.some((t) => t.serverPath === destination);

  return (
    <StepFrame
      title="Upload document"
      onBack={() => goTo(9)}
      backLabel="Results"
      primary={
        <Button
          onClick={() => {
            setSubmitted(true);
            if (!fileProblem && destinationValid) goTo(11);
          }}
        >
          Review upload <ArrowRight aria-hidden="true" />
        </Button>
      }
    >
      <div className="space-y-3">
        <h2 className="text-sm font-semibold">File</h2>
        <input
          ref={inputRef}
          id="file-input"
          type="file"
          className="sr-only"
          tabIndex={-1}
          aria-hidden="true"
          onChange={(e) => {
            setFile(e.target.files?.[0] ?? null);
            e.target.value = "";
          }}
        />
        <div className="flex flex-wrap items-center gap-4 rounded-lg border border-dashed p-4">
          <Button
            variant="outline"
            onClick={() => inputRef.current?.click()}
            aria-describedby={
              submitted && fileProblem ? "file-error" : undefined
            }
          >
            <FileUp aria-hidden="true" />{" "}
            {file ? "Choose another file" : "Choose file"}
          </Button>
          {file ? (
            <dl className="grid min-w-0 flex-1 grid-cols-[5rem_minmax(0,1fr)] gap-x-3 gap-y-0.5 text-sm">
              <dt className="text-muted-foreground">Name</dt>
              <dd className="truncate font-medium" title={file.name}>
                {file.name}
              </dd>
              <dt className="text-muted-foreground">Type</dt>
              <dd className="truncate">{fileTypeLabel(file)}</dd>
              <dt className="text-muted-foreground">Size</dt>
              <dd className="tabular-nums">{formatBytes(file.size)}</dd>
            </dl>
          ) : (
            <p className="text-sm">No file selected.</p>
          )}
        </div>
        {(submitted || file) && fileProblem && (
          <p id="file-error" className="text-sm text-red-700 dark:text-red-300">
            {fileProblem}
          </p>
        )}
      </div>

      <div className="space-y-3">
        <h2 id="dest-label" className="text-sm font-semibold">
          Destination folder
        </h2>
        {targets.length === 0 ? (
          <p className="text-sm">
            No project folder is available. Go back and create the structure
            first.
          </p>
        ) : (
          <RadioGroup
            aria-labelledby="dest-label"
            value={destination ?? ""}
            onValueChange={setDestination}
            className="grid gap-0 divide-y rounded-lg border sm:grid-cols-2 sm:divide-y-0"
          >
            {targets.map((t) => (
              <label
                key={t.serverPath}
                htmlFor={`dest-${t.serverPath}`}
                className="flex cursor-pointer items-center gap-3 border-b px-3 py-2.5 hover:bg-accent has-[[data-state=checked]]:bg-accent sm:odd:border-r"
              >
                <RadioGroupItem
                  id={`dest-${t.serverPath}`}
                  value={t.serverPath}
                />
                <span className="min-w-0 flex-1 truncate font-mono text-sm">
                  {t.name}
                </span>
              </label>
            ))}
          </RadioGroup>
        )}
        {submitted && !destinationValid && targets.length > 0 && (
          <p className="text-sm text-red-700 dark:text-red-300">
            Select a destination folder.
          </p>
        )}
      </div>
    </StepFrame>
  );
}

const STAGE_LABEL: Record<UploadStage, string> = {
  reading: "Reading the file",
  checking: "Checking the destination for an existing file",
  uploading: "Uploading to SharePoint",
  verifying: "Verifying the file in SharePoint",
  done: "Done",
};
const STAGE_PCT: Record<UploadStage, number> = {
  reading: 10,
  checking: 25,
  uploading: 60,
  verifying: 90,
  done: 100,
};

// ---------- 11. Upload confirmation ----------
export function UploadConfirmStep() {
  const config = useWizard((s) => s.config);
  const file = useWizard((s) => s.file);
  const destination = useWizard((s) => s.destinationServerPath);
  const targetFileName = useWizard((s) => s.targetFileName);
  const setTargetFileName = useWizard((s) => s.setTargetFileName);
  const setUploadResult = useWizard((s) => s.setUploadResult);
  const busy = useWizard((s) => s.busy);
  const setBusy = useWizard((s) => s.setBusy);
  const goTo = useWizard((s) => s.goTo);
  const targets = useUploadTargets();
  const folder = targets.find((t) => t.serverPath === destination);
  const [confirmed, setConfirmed] = useState(false);
  const [stage, setStage] = useState<UploadStage | null>(null);
  const [renaming, setRenaming] = useState(false);
  const [renameError, setRenameError] = useState<unknown>(null);

  const name = targetFileName ?? file?.name ?? "";
  const existing = useQuery({
    queryKey: ["file-exists", config?.site.url, destination, name],
    queryFn: async () => {
      try {
        return await SharePointFileService.fileExists(
          config!.site.url,
          fileServerPath(destination!, name),
        );
      } catch (e) {
        throw toAppError(e, {
          operation: "GetFileMetadataByPath",
          action: "check the destination for an existing file",
        });
      }
    },
    enabled: Boolean(config && destination && name && file),
    retry: false,
    staleTime: 0,
  });

  if (!config || !file || !folder) {
    return (
      <StepFrame title="Upload confirmation" onBack={() => goTo(10)}>
        <p className="text-sm">
          {!file
            ? "No file selected. Choose the file again — files are not kept after a reload."
            : "Select a destination folder first."}
        </p>
      </StepFrame>
    );
  }

  const fullParts = [
    config.site.label,
    config.library.displayName,
    ...splitRel(folder.relativePath),
    name,
  ];

  async function chooseNewName() {
    setRenaming(true);
    setRenameError(null);
    try {
      const suggestion = await SharePointFileService.suggestAvailableName(
        config!.site.url,
        folder!.serverPath,
        file!.name,
      );
      setTargetFileName(suggestion);
      setConfirmed(false);
    } catch (e) {
      setRenameError(
        toAppError(e, {
          operation: "GetFileMetadataByPath",
          action: "find a free file name",
        }),
      );
    } finally {
      setRenaming(false);
    }
  }

  async function upload() {
    setBusy(true);
    try {
      const result = await SharePointFileService.upload(
        config!.site.url,
        {
          file: file!,
          folderServerPath: folder!.serverPath,
          folderRelativePath: folder!.relativePath,
          targetFileName: name,
        },
        setStage,
      );
      setUploadResult(result);
      void queryClient.invalidateQueries({ queryKey: ["contents"] });
      if (result.duplicate) {
        // Someone added a file with this name since the check: stay here and show the duplicate choice.
        await existing.refetch();
        setStage(null);
        setConfirmed(false);
        return;
      }
      goTo(12);
    } finally {
      setBusy(false);
    }
  }

  const duplicate = existing.data === true;
  const canUpload = confirmed && existing.data === false && !busy;

  return (
    <StepFrame
      title="Upload confirmation"
      onBack={busy ? undefined : () => goTo(10)}
      primary={
        <Button disabled={!canUpload} onClick={() => void upload()}>
          {busy ? <Spinner /> : <Upload aria-hidden="true" />} Upload file
        </Button>
      }
    >
      <dl className="grid gap-x-6 gap-y-3 rounded-lg border bg-card p-4 text-sm sm:grid-cols-[10rem_minmax(0,1fr)]">
        <dt className="text-muted-foreground">Local file</dt>
        <dd className="min-w-0 break-all">{file.name}</dd>
        <dt className="text-muted-foreground">Type</dt>
        <dd>{fileTypeLabel(file)}</dd>
        <dt className="text-muted-foreground">Size</dt>
        <dd className="tabular-nums">{formatBytes(file.size)}</dd>
        <dt className="text-muted-foreground">Name in SharePoint</dt>
        <dd className="min-w-0 font-medium break-all">{name}</dd>
        <dt className="text-muted-foreground">Destination</dt>
        <dd className="min-w-0">
          <PathLine parts={fullParts} />
        </dd>
        <dt className="text-muted-foreground">Site</dt>
        <dd className="min-w-0 break-all">{config.site.url}</dd>
      </dl>

      {existing.isPending && (
        <p role="status" className="flex items-center gap-2 text-sm">
          <Spinner /> Checking for a file with the same name…
        </p>
      )}
      {existing.isError && (
        <ErrorPanel
          title="The destination couldn't be checked"
          error={existing.error}
          action={
            <Button
              size="sm"
              variant="outline"
              onClick={() => existing.refetch()}
            >
              Try again
            </Button>
          }
        />
      )}
      {duplicate && (
        <div
          role="alert"
          className="rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm text-amber-950 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100"
        >
          <div className="flex gap-3">
            <AlertTriangle
              className="mt-0.5 size-4 shrink-0"
              aria-hidden="true"
            />
            <div className="space-y-3">
              <p>
                <span className="font-medium">Duplicate filename.</span> "{name}
                " already exists in {folder.name}. This app never overwrites
                files.
              </p>
              <div className="flex flex-wrap gap-2">
                <Button
                  size="sm"
                  onClick={() => void chooseNewName()}
                  disabled={renaming}
                >
                  {renaming && <Spinner />} Upload under a new name
                </Button>
                <Button size="sm" variant="outline" onClick={() => goTo(10)}>
                  Choose another file or folder
                </Button>
              </div>
              {renameError !== null && (
                <ErrorPanel title="No free name found" error={renameError} />
              )}
            </div>
          </div>
        </div>
      )}

      {existing.data === false && (
        <label
          htmlFor="confirm-upload"
          className="flex items-start gap-3 rounded-lg border p-4"
        >
          <Checkbox
            id="confirm-upload"
            checked={confirmed}
            onCheckedChange={(v) => setConfirmed(v === true)}
            disabled={busy}
            className="mt-0.5"
          />
          <span className="text-sm">
            I want to upload this file to the destination shown above.
          </span>
        </label>
      )}

      {stage && (
        <div aria-live="polite" className="space-y-2">
          <div className="flex justify-between text-sm">
            <span>{STAGE_LABEL[stage]}…</span>
            <span className="tabular-nums">{STAGE_PCT[stage]}%</span>
          </div>
          <div
            className="h-2 overflow-hidden rounded-full bg-muted"
            role="progressbar"
            aria-label="Upload progress"
            aria-valuemin={0}
            aria-valuemax={100}
            aria-valuenow={STAGE_PCT[stage]}
          >
            <div
              className="h-full rounded-full bg-primary transition-[width] duration-500"
              style={{ width: `${STAGE_PCT[stage]}%` }}
            />
          </div>
        </div>
      )}
    </StepFrame>
  );
}

// ---------- 12. Upload result and completion ----------
export function UploadResultStep() {
  const config = useWizard((s) => s.config);
  const result = useWizard((s) => s.uploadResult);
  const folderResults = useWizard((s) => s.folderResults);
  const startNewUpload = useWizard((s) => s.startNewUpload);
  const startNewProject = useWizard((s) => s.startNewProject);
  const goTo = useWizard((s) => s.goTo);
  if (!config || !result) {
    return (
      <StepFrame title="Upload result" onBack={() => goTo(10)}>
        <p className="text-sm">No upload has run yet.</p>
      </StepFrame>
    );
  }
  const ok = result.status === "Uploaded";
  const pending = result.status === "Pending";
  const folderPath = result.serverPath.slice(
    0,
    result.serverPath.lastIndexOf("/"),
  );
  const root = folderResults?.[0];
  const created = (folderResults ?? []).filter(
    (r) => r?.status === "Created",
  ).length;
  const existed = (folderResults ?? []).filter(
    (r) => r?.status === "Already exists",
  ).length;

  return (
    <StepFrame
      title="Upload result and completion"
      secondary={
        <Button variant="outline" onClick={startNewUpload}>
          <Upload aria-hidden="true" />{" "}
          {ok ? "Upload another document" : "Try again"}
        </Button>
      }
      primary={
        <Button onClick={startNewProject}>
          <FolderPlus aria-hidden="true" /> Start new project
        </Button>
      }
    >
      <div
        role={ok || pending ? "status" : "alert"}
        className={
          ok
            ? "flex gap-3 rounded-lg border border-green-200 bg-green-50 p-4 text-green-950 dark:border-green-800 dark:bg-green-950 dark:text-green-100"
            : pending
              ? "flex gap-3 rounded-lg border border-amber-200 bg-amber-50 p-4 text-amber-950 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100"
              : "flex gap-3 rounded-lg border border-red-200 bg-red-50 p-4 text-red-950 dark:border-red-800 dark:bg-red-950 dark:text-red-100"
        }
      >
        {ok ? (
          <CheckCircle2 className="mt-0.5 size-5 shrink-0" aria-hidden="true" />
        ) : (
          <XCircle className="mt-0.5 size-5 shrink-0" aria-hidden="true" />
        )}
        <div className="space-y-1 text-sm">
          <p className="flex flex-wrap items-center gap-2 font-medium">
            {ok
              ? "Upload complete"
              : pending
                ? "Upload pending"
                : result.status === "Blocked"
                  ? "Upload blocked"
                  : "Upload failed"}{" "}
            <StatusBadge status={result.status} />
          </p>
          <p>{result.message}</p>
          {result.diagnostic && (
            <p className="font-mono text-xs">Support: {result.diagnostic}</p>
          )}
        </div>
      </div>

      <dl className="grid gap-x-6 gap-y-3 rounded-lg border bg-card p-4 text-sm sm:grid-cols-[10rem_minmax(0,1fr)]">
        <dt className="text-muted-foreground">File</dt>
        <dd className="min-w-0 font-medium break-all">{result.fileName}</dd>
        <dt className="text-muted-foreground">Size</dt>
        <dd className="tabular-nums">{formatBytes(result.sizeBytes)}</dd>
        <dt className="text-muted-foreground">SharePoint location</dt>
        <dd className="min-w-0 break-all">
          {config.site.url}
          {result.serverPath}
        </dd>
        {root && (
          <>
            <dt className="text-muted-foreground">Project structure</dt>
            <dd>
              {root.name} · {created} created, {existed} already existed
            </dd>
          </>
        )}
      </dl>

      <div className="flex flex-wrap gap-4">
        {(ok || pending) && (
          <a
            href={safeUrl(sharePointLink(config.site.url, folderPath)) ?? "#"}
            target="_blank"
            rel="noreferrer"
            className="inline-flex items-center gap-1.5 text-sm font-medium text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring"
          >
            Open folder in SharePoint{" "}
            <ExternalLink className="size-3.5" aria-hidden="true" />
          </a>
        )}
        {root && (
          <a
            href={
              safeUrl(sharePointLink(config.site.url, root.serverPath)) ?? "#"
            }
            target="_blank"
            rel="noreferrer"
            className="inline-flex items-center gap-1.5 text-sm font-medium text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring"
          >
            Open project folder{" "}
            <ExternalLink className="size-3.5" aria-hidden="true" />
          </a>
        )}
      </div>
    </StepFrame>
  );
}
