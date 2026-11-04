// Upload hand-off to the "SPSB Upload" Power Automate flow.
//
// Why: the app's SharePoint connection sends request bodies as UTF-8 text, so
// binary files (PDF, Office, images) can't be written directly. Text CAN be
// written, so the app writes an upload request — JSON with the file content as
// base64 — into a queue folder. The flow (triggered by new files in that folder)
// decodes it, writes the real file through Microsoft Graph using an app
// registration whose secret/certificate stays in Azure Key Vault, and writes a
// result file back. No flow URL, secret or token is ever held by the app.
//
// Contract (must match FLOW-SETUP.md):
//   <queue>/<requestId>.request.json  — written by the app
//   <queue>/<requestId>.result.json   — written by the flow: {"status":"Uploaded"|"Failed","message":"..."}
import { SharePointService } from "../../generated/services/SharePointService";
import type { SharePointConfiguration, UploadResult } from "@/models";
import { statusOf, toAppError } from "./errors";
import { joinPath } from "./validation";

export const UPLOAD_REQUEST_FORMAT = "spsb.upload-request";
export const DEFAULT_QUEUE_FOLDER = "_UploadQueue";
const POLL_INTERVAL_MS = 4000;
const MAX_WAIT_MS = 3 * 60 * 1000;

export interface UploadRequest {
  format: typeof UPLOAD_REQUEST_FORMAT;
  version: 1;
  requestId: string;
  requestedAt: string;
  siteUrl: string;
  /** Library (list) id used by the flow to resolve the Graph drive. */
  libraryId: string;
  /** Destination folder relative to the library root, e.g. "Projects/P-1 - X/02. Contracten". */
  targetFolder: string;
  fileName: string;
  contentType: string;
  size: number;
  contentBase64: string;
}

export function queueServerPath(config: SharePointConfiguration): string {
  return `/${joinPath(config.library.rootPath, config.uploadFlow?.queueFolder ?? DEFAULT_QUEUE_FOLDER)}`;
}

/** Folder path relative to the library root, from a server-relative path. */
export function libraryRelativePath(config: SharePointConfiguration, serverPath: string): string {
  const root = config.library.rootPath.replace(/\/+$/, "");
  const p = serverPath.replace(/\/+$/, "");
  return p.toLowerCase().startsWith(root.toLowerCase()) ? joinPath(p.slice(root.length)) : joinPath(p);
}

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(binary);
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Reads the flow's result file; null while it doesn't exist yet. Tolerates text, JSON object or a $content envelope. */
async function readResult(siteUrl: string, path: string): Promise<{ status: string; message?: string } | null> {
  const r = await SharePointService.GetFileContentByPath(siteUrl, path, false);
  if (!r.success) {
    if (statusOf(r.error) === 404) return null;
    throw toAppError(r.error, { operation: "GetFileContentByPath", action: "read the upload flow's result" });
  }
  let payload: unknown = r.data;
  if (payload && typeof payload === "object" && "$content" in payload) {
    try {
      payload = atob(String((payload as Record<string, unknown>)["$content"]));
    } catch {
      payload = null;
    }
  }
  if (typeof payload === "string") {
    const text = payload;
    try {
      payload = JSON.parse(text);
    } catch {
      try {
        payload = JSON.parse(atob(text));
      } catch {
        payload = null;
      }
    }
  }
  if (payload && typeof payload === "object" && typeof (payload as { status?: unknown }).status === "string") {
    const p = payload as { status: string; message?: unknown };
    return { status: p.status, message: typeof p.message === "string" ? p.message : undefined };
  }
  return { status: "Unreadable" };
}

export const UploadFlowService = {
  /** Creates the queue folder when missing (idempotent). */
  async ensureQueueFolder(config: SharePointConfiguration): Promise<void> {
    const path = queueServerPath(config);
    const exists = await SharePointService.GetFolderMetadataByPath(config.site.url, path);
    if (exists.success) return;
    if (statusOf(exists.error) !== 404) throw toAppError(exists.error, { operation: "GetFolderMetadataByPath", action: "open the upload queue folder" });
    const created = await SharePointService.CreateNewFolder(config.site.url, config.library.id, {
      path: config.uploadFlow?.queueFolder ?? DEFAULT_QUEUE_FOLDER,
    });
    if (!created.success) throw toAppError(created.error, { operation: "CreateNewFolder", action: "create the upload queue folder" });
  },

  /**
   * Hands the file to the flow and waits for its result. The caller has already
   * checked that the target name is free. Success is reported only after the
   * flow reports success AND the file is found at the destination with the same size.
   */
  async upload(
    config: SharePointConfiguration,
    bytes: Uint8Array,
    file: File,
    targetFolderServerPath: string,
    fileName: string,
    onStage: (s: "handing-over" | "processing" | "verifying") => void,
  ): Promise<UploadResult> {
    const serverPath = `${targetFolderServerPath.replace(/\/+$/, "")}/${fileName}`;
    const base = { fileName, serverPath, sizeBytes: file.size };
    const site = config.site.url;
    const requestId = crypto.randomUUID();
    const queue = queueServerPath(config);

    onStage("handing-over");
    await this.ensureQueueFolder(config);
    const request: UploadRequest = {
      format: UPLOAD_REQUEST_FORMAT,
      version: 1,
      requestId,
      requestedAt: new Date().toISOString(),
      siteUrl: site,
      libraryId: config.library.id,
      targetFolder: libraryRelativePath(config, targetFolderServerPath),
      fileName,
      contentType: file.type || "application/octet-stream",
      size: file.size,
      contentBase64: toBase64(bytes),
    };
    const created = await SharePointService.CreateFile(site, queue, `${requestId}.request.json`, JSON.stringify(request));
    if (!created.success) {
      const err = toAppError(created.error, { operation: "CreateFile", action: "hand the file to the upload flow" });
      return { ...base, status: "Failed", message: err.userMessage, diagnostic: err.diagnostic };
    }

    onStage("processing");
    const resultPath = `${queue}/${requestId}.result.json`;
    const started = Date.now();
    let result: { status: string; message?: string } | null = null;
    while (Date.now() - started < MAX_WAIT_MS) {
      await sleep(POLL_INTERVAL_MS);
      result = await readResult(site, resultPath);
      if (result) break;
    }
    if (result) {
      // Best effort: remove the result file the flow wrote for this request.
      const meta = await SharePointService.GetFileMetadataByPath(site, resultPath);
      if (meta.success && meta.data?.Id) await SharePointService.DeleteFile(site, meta.data.Id);
    }
    if (!result) {
      return {
        ...base,
        status: "Pending",
        message:
          "The file was handed to the upload flow, but it hasn't finished yet. Open the folder in SharePoint in a few minutes to check. If the file doesn't appear, ask your administrator to check the upload flow's run history.",
        diagnostic: `request ${requestId}: no result file after ${MAX_WAIT_MS / 1000}s`,
      };
    }
    if (result.status !== "Uploaded") {
      return {
        ...base,
        status: "Failed",
        message: result.message ? `The upload flow reported: ${result.message}` : "The upload flow couldn't store the file.",
        diagnostic: `request ${requestId}: flow status ${result.status}`,
      };
    }

    onStage("verifying");
    const meta = await SharePointService.GetFileMetadataByPath(site, serverPath);
    if (!meta.success) {
      const err = toAppError(meta.error, { operation: "GetFileMetadataByPath", action: "verify the uploaded file" });
      return { ...base, status: "Failed", message: `The flow reported success, but the file couldn't be found at the destination. ${err.userMessage}`, diagnostic: `request ${requestId}; ${err.diagnostic}` };
    }
    const stored = meta.data?.Size;
    if (typeof stored === "number" && stored !== file.size) {
      return {
        ...base,
        status: "Failed",
        message: "The file in SharePoint differs from your original. Check it in SharePoint and ask your administrator to review the upload flow.",
        diagnostic: `request ${requestId}: size local ${file.size} B, stored ${stored} B`,
      };
    }
    return { ...base, status: "Uploaded", message: "File uploaded through the upload flow. SharePoint confirms the file and its size." };
  },
};
