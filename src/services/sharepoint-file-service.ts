// SharePointFileService — duplicate detection and verified upload.
// Never overwrites: the destination name is checked right before the upload and
// the upload is refused when a file with that name exists.
import type { UploadInput, UploadResult } from "@/models";
import { AppError, statusOf, toAppError } from "./errors";
import { GraphItemNotFoundError, GraphSharePoint } from "./graph-sharepoint";
import { MAX_UPLOAD_BYTES, formatBytes, validateItemName } from "./validation";

export type UploadStage = "reading" | "checking" | "uploading" | "verifying" | "done";

export function fileServerPath(
  folderServerPath: string,
  fileName: string,
): string {
  return `${folderServerPath.replace(/\/+$/, "")}/${fileName}`;
}

/** Validates the local file before anything is sent. Returns a user message or null. */
export function validateLocalFile(file: File | null): string | null {
  if (!file) return "Select a file to upload.";
  if (file.size === 0) return "The selected file is empty.";
  if (file.size > MAX_UPLOAD_BYTES)
    return `The file is ${formatBytes(file.size)}. This app uploads files up to ${formatBytes(MAX_UPLOAD_BYTES)}.`;
  return validateItemName(file.name, "file");
}

function splitName(name: string): { stem: string; ext: string } {
  const dot = name.lastIndexOf(".");
  return dot > 0
    ? { stem: name.slice(0, dot), ext: name.slice(dot) }
    : { stem: name, ext: "" };
}

export const SharePointFileService = {
  async fileExists(siteUrl: string, serverPath: string): Promise<boolean> {
    try {
      await GraphSharePoint.itemAtPath(siteUrl, serverPath);
      return true;
    } catch (error) {
      if (error instanceof GraphItemNotFoundError) return false;
      throw toAppError(error, { operation: "GetFileMetadataByPath", action: "check the destination for an existing file" });
    }
  },

  /** First free "name (n).ext" in the folder. */
  async suggestAvailableName(
    siteUrl: string,
    folderServerPath: string,
    fileName: string,
  ): Promise<string> {
    const { stem, ext } = splitName(fileName);
    for (let n = 2; n <= 50; n++) {
      const candidate = `${stem} (${n})${ext}`;
      if (
        !(await this.fileExists(
          siteUrl,
          fileServerPath(folderServerPath, candidate),
        ))
      )
        return candidate;
    }
    throw new AppError(
      "conflict",
      "Too many files with this name already exist in the folder. Rename the file on your computer and try again.",
      "suggestAvailableName exhausted 50 candidates",
      409,
    );
  },

  async upload(siteUrl: string, input: UploadInput, onStage: (s: UploadStage) => void): Promise<UploadResult> {
    const serverPath = fileServerPath(input.folderServerPath, input.targetFileName);
    const baseResult = { fileName: input.targetFileName, serverPath, sizeBytes: input.file.size };

    const localProblem =
      validateLocalFile(input.file) ??
      validateItemName(input.targetFileName, "file");
    if (localProblem)
      return { ...baseResult, status: "Blocked", message: localProblem };

    try {
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
      const created = await GraphSharePoint.upload(siteUrl, serverPath, input.file);

      onStage("verifying");
      const meta = await GraphSharePoint.itemAtPath(siteUrl, serverPath);
      const storedSize = meta.size;
      if (meta.id !== created.id || (typeof storedSize === "number" && storedSize !== input.file.size)) {
        return {
          ...baseResult,
          status: "Failed",
          message: "The uploaded file could not be verified. Check the destination in SharePoint before trying again.",
          diagnostic: "Uploaded file identity or size did not match the verified item.",
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
      if (statusOf(e) === 409) {
        return { ...baseResult, status: "Blocked", duplicate: true, message: "A file with this name already exists. Nothing was overwritten." };
      }
      const err = toAppError(e, { operation: "CreateFile", action: "upload the file" });
      return { ...baseResult, status: "Failed", message: err.userMessage, diagnostic: err.diagnostic };
    }
  },
};
