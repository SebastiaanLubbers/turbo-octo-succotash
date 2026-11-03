import { useRef, useState } from "react";
import { AlertTriangle, CheckCircle2, ExternalLink, FileUp, RotateCw, Upload, X } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Spinner } from "@/components/ui/spinner";
import { useConfirm } from "@/hooks/use-confirm";
import { safeUrl } from "@/lib/safe-url";
import type { SharePointConfiguration, UploadResult } from "@/models";
import { toAppError } from "@/services/errors";
import {
  BINARY_NOT_SUPPORTED_MESSAGE,
  SharePointFileService,
  canUploadInApp,
  fileServerPath,
  validateLocalFile,
  type UploadStage,
} from "@/services/sharepoint-file-service";
import { formatBytes } from "@/services/validation";
import { sharePointLink } from "@/state/selectors";
import { ErrorPanel } from "@/components/wizard/error-panel";
import { PathLine } from "@/components/wizard/destination-tree";

const STAGE_LABEL: Record<UploadStage, string> = {
  reading: "Reading the file",
  checking: "Checking for an existing file",
  uploading: "Uploading to SharePoint",
  verifying: "Verifying the file in SharePoint",
  done: "Done",
};

/**
 * Adds one local file to an existing SharePoint folder.
 * Same rules as the wizard: text files only, duplicate check, explicit
 * confirmation, verified upload, never overwrite.
 */
export function AddFilePanel({
  config,
  folderServerPath,
  pathParts,
  onClose,
  onUploaded,
}: {
  config: SharePointConfiguration;
  folderServerPath: string;
  pathParts: string[];
  onClose: () => void;
  onUploaded: () => void;
}) {
  const confirm = useConfirm();
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [binary, setBinary] = useState<boolean | null>(null);
  const [stage, setStage] = useState<UploadStage | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<unknown>(null);
  const [result, setResult] = useState<UploadResult | null>(null);
  const problem = validateLocalFile(file);
  const folderLink = safeUrl(sharePointLink(config.site.url, folderServerPath)) ?? "#";

  async function pick(f: File | null) {
    setFile(f);
    setResult(null);
    setError(null);
    setBinary(null);
    if (f && !validateLocalFile(f)) setBinary(!(await canUploadInApp(f)));
  }

  async function upload() {
    if (!file) return;
    setBusy(true);
    setError(null);
    setResult(null);
    try {
      let name = file.name;
      setStage("checking");
      if (await SharePointFileService.fileExists(config.site.url, fileServerPath(folderServerPath, name))) {
        const suggestion = await SharePointFileService.suggestAvailableName(config.site.url, folderServerPath, name);
        const ok = await confirm({
          title: "Duplicate filename",
          description: `"${name}" already exists in this folder and won't be overwritten. Upload it as "${suggestion}" instead?`,
          confirmLabel: `Upload as ${suggestion}`,
        });
        if (!ok) return;
        name = suggestion;
      } else {
        const ok = await confirm({
          title: "Upload file?",
          description: `"${name}" (${formatBytes(file.size)}) will be uploaded to ${[...pathParts].join(" › ")}.`,
          confirmLabel: "Upload file",
        });
        if (!ok) return;
      }
      const r = await SharePointFileService.upload(
        config.site.url,
        { file, folderServerPath, folderRelativePath: "", targetFileName: name },
        setStage,
      );
      setResult(r);
      if (r.status === "Uploaded") {
        toast.success(`${r.fileName} uploaded`);
        onUploaded();
      }
    } catch (e) {
      setError(toAppError(e, { operation: "GetFileMetadataByPath", action: "check the destination" }));
    } finally {
      setBusy(false);
      setStage(null);
    }
  }

  return (
    <section aria-labelledby="add-file-title" className="space-y-4 rounded-lg border bg-card p-4">
      <div className="flex items-start gap-3">
        <div className="min-w-0 flex-1">
          <h2 id="add-file-title" className="font-semibold">Add file</h2>
          <div className="mt-1"><PathLine parts={pathParts} /></div>
        </div>
        <Button variant="ghost" size="icon" onClick={onClose} aria-label="Close add file" disabled={busy}>
          <X aria-hidden="true" />
        </Button>
      </div>

      <input
        ref={inputRef}
        type="file"
        className="sr-only"
        tabIndex={-1}
        aria-hidden="true"
        onChange={(e) => {
          void pick(e.target.files?.[0] ?? null);
          e.target.value = "";
        }}
      />
      <div className="flex flex-wrap items-center gap-4">
        <Button variant="outline" onClick={() => inputRef.current?.click()} disabled={busy}>
          <FileUp aria-hidden="true" /> {file ? "Choose another file" : "Choose file"}
        </Button>
        {file ? (
          <p className="min-w-0 text-sm">
            <span className="font-medium break-all">{file.name}</span>
            <span className="text-muted-foreground"> · {file.type || "unknown type"} · {formatBytes(file.size)}</span>
          </p>
        ) : (
          <p className="text-sm">No file selected.</p>
        )}
      </div>
      {file && problem && <p className="text-sm text-red-700 dark:text-red-300">{problem}</p>}

      {binary === true && (
        <div role="alert" className="flex gap-3 rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm text-amber-950 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
          <div className="space-y-2">
            <p>{BINARY_NOT_SUPPORTED_MESSAGE}</p>
            <div className="flex flex-wrap gap-2">
              <Button size="sm" asChild>
                <a href={folderLink} target="_blank" rel="noreferrer">
                  Open folder in SharePoint <ExternalLink aria-hidden="true" />
                  <span className="sr-only">(opens in a new tab)</span>
                </a>
              </Button>
              <Button size="sm" variant="outline" onClick={onUploaded}>
                <RotateCw aria-hidden="true" /> Refresh folder
              </Button>
            </div>
          </div>
        </div>
      )}

      {stage && (
        <p role="status" className="flex items-center gap-2 text-sm"><Spinner /> {STAGE_LABEL[stage]}…</p>
      )}
      {error !== null && <ErrorPanel error={error} />}
      {result && result.status === "Uploaded" && (
        <p role="status" className="flex items-center gap-2 text-sm text-green-800 dark:text-green-200">
          <CheckCircle2 className="size-4" aria-hidden="true" /> {result.fileName} uploaded. {result.message}
        </p>
      )}
      {result && result.status !== "Uploaded" && (
        <div role="alert" className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-900 dark:border-red-800 dark:bg-red-950 dark:text-red-100">
          <p>{result.message}</p>
          {result.diagnostic && <p className="mt-1 font-mono text-xs">Support: {result.diagnostic}</p>}
        </div>
      )}

      <div className="flex justify-end">
        <Button onClick={() => void upload()} disabled={!file || Boolean(problem) || binary !== false || busy}>
          {busy ? <Spinner /> : <Upload aria-hidden="true" />} Upload
        </Button>
      </div>
    </section>
  );
}
