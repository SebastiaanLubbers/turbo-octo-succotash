// SharePointFileService — duplicate detection and verified upload.
// Never overwrites: the destination name is checked right before the upload and
// the upload is refused when a file with that name exists.
import { SharePointService } from "../../generated/services/SharePointService";
import type { UploadInput, UploadResult } from "@/models";
import { AppError, statusOf, toAppError } from "./errors";
import { MAX_UPLOAD_BYTES, formatBytes, validateItemName } from "./validation";

export type UploadStage = "reading" | "checking" | "uploading" | "verifying" | "done";

export function fileServerPath(folderServerPath: string, fileName: string): string {
  return `${folderServerPath.replace(/\/+$/, "")}/${fileName}`;
}

/** Validates the local file before anything is sent. Returns a user message or null. */
export function validateLocalFile(file: File | null): string | null {
  if (!file) return "Select a file to upload.";
  if (file.size === 0) return "The selected file is empty.";
  if (file.size > MAX_UPLOAD_BYTES) return `The file is ${formatBytes(file.size)}. This app uploads files up to ${formatBytes(MAX_UPLOAD_BYTES)}.`;
  return validateItemName(file.name, "file");
}

function splitName(name: string): { stem: string; ext: string } {
  const dot = name.lastIndexOf(".");
  return dot > 0 ? { stem: name.slice(0, dot), ext: name.slice(dot) } : { stem: name, ext: "" };
}

/** Exact UTF-8 text when the bytes are valid UTF-8; otherwise null (binary file). */
function asExactText(bytes: Uint8Array): string | null {
  try {
    return new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch {
    return null;
  }
}

/**
 * The SharePoint connection available to this app sends request bodies as
 * UTF-8 text, so only files whose bytes are valid UTF-8 arrive unchanged.
 * Binary files (PDF, Office documents, images, archives) are blocked before
 * any SharePoint call instead of being stored in a damaged form.
 */
export async function canUploadInApp(file: File): Promise<boolean> {
  return asExactText(new Uint8Array(await file.arrayBuffer())) !== null;
}

export const BINARY_NOT_SUPPORTED_MESSAGE =
  "This file type can't be uploaded from the app: the SharePoint connection available to it only transfers text files (such as TXT, CSV, JSON or XML). Open the destination folder in SharePoint and upload the file there.";

export const SharePointFileService = {
  async fileExists(siteUrl: string, serverPath: string): Promise<boolean> {
    const result = await SharePointService.GetFileMetadataByPath(siteUrl, serverPath);
    if (result.success) return true;
    if (statusOf(result.error) === 404) return false;
    throw toAppError(result.error, { operation: "GetFileMetadataByPath", action: "check the destination for an existing file" });
  },

  /** First free "name (n).ext" in the folder. */
  async suggestAvailableName(siteUrl: string, folderServerPath: string, fileName: string): Promise<string> {
    const { stem, ext } = splitName(fileName);
    for (let n = 2; n <= 50; n++) {
      const candidate = `${stem} (${n})${ext}`;
      if (!(await this.fileExists(siteUrl, fileServerPath(folderServerPath, candidate)))) return candidate;
    }
    throw new AppError("conflict", "Too many files with this name already exist in the folder. Rename the file on your computer and try again.", "suggestAvailableName exhausted 50 candidates", 409);
  },

  async upload(siteUrl: string, input: UploadInput, onStage: (s: UploadStage) => void): Promise<UploadResult> {
    const serverPath = fileServerPath(input.folderServerPath, input.targetFileName);
    const baseResult = { fileName: input.targetFileName, serverPath, sizeBytes: input.file.size };

    const localProblem = validateLocalFile(input.file) ?? validateItemName(input.targetFileName, "file");
    if (localProblem) return { ...baseResult, status: "Blocked", message: localProblem };

    try {
      onStage("reading");
      const text = asExactText(new Uint8Array(await input.file.arrayBuffer()));
      if (text === null) {
        return { ...baseResult, status: "Blocked", binary: true, message: BINARY_NOT_SUPPORTED_MESSAGE };
      }

      onStage("checking");
      if (await this.fileExists(siteUrl, serverPath)) {
        return {
          ...baseResult,
          status: "Blocked",
          duplicate: true,
          message: `A file named "${input.targetFileName}" already exists in this folder. Nothing was overwritten. Choose a new name or another folder.`,
        };
      }

      onStage("uploading");
      const created = await SharePointService.CreateFile(siteUrl, input.folderServerPath, input.targetFileName, text);
      if (!created.success) {
        const err = toAppError(created.error, { operation: "CreateFile", action: "upload the file" });
        return { ...baseResult, status: "Failed", message: err.userMessage, diagnostic: err.diagnostic };
      }

      onStage("verifying");
      const meta = await SharePointService.GetFileMetadataByPath(siteUrl, serverPath);
      if (!meta.success) {
        const err = toAppError(meta.error, { operation: "GetFileMetadataByPath", action: "verify the uploaded file" });
        return { ...baseResult, status: "Failed", message: `The upload couldn't be verified. ${err.userMessage}`, diagnostic: err.diagnostic };
      }
      const storedSize = meta.data?.Size;
      if (typeof storedSize === "number" && storedSize !== input.file.size) {
        // The app created this file moments ago under a name that was free, so removing it is safe.
        const fileId = meta.data?.Id ?? created.data?.Id;
        const removed = fileId ? (await SharePointService.DeleteFile(siteUrl, fileId)).success : false;
        return {
          ...baseResult,
          status: "Failed",
          message: removed
            ? "SharePoint stored a file that differs from your original, so the app removed it again. Upload this file directly in SharePoint."
            : "SharePoint stored a file that differs from your original. Delete it in SharePoint and upload the file there directly.",
          diagnostic: `Size mismatch: local ${input.file.size} B, stored ${storedSize} B; cleanup ${removed ? "succeeded" : "failed"}`,
        };
      }
      onStage("done");
      return {
        ...baseResult,
        status: "Uploaded",
        message:
          typeof storedSize === "number"
            ? "File uploaded. SharePoint confirms the file and its size."
            : "File uploaded. SharePoint confirms the file exists at this location.",
      };
    } catch (e) {
      const err = toAppError(e, { operation: "CreateFile", action: "upload the file" });
      return { ...baseResult, status: "Failed", message: err.userMessage, diagnostic: err.diagnostic };
    }
  },
};
