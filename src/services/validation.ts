// Validation rules applied before any SharePoint mutation.
import type { ProjectInput } from "@/models";

const ILLEGAL_CHARS = /["*:<>?/\\|#%]/;
const RESERVED_NAMES = new Set([
  "con", "prn", "aux", "nul",
  "com0", "com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8", "com9",
  "lpt0", "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
  "_vti_", "desktop.ini", "forms",
]);
export const MAX_SEGMENT_LENGTH = 255;
/** SharePoint Online limits the full decoded path to 400 characters. */
export const MAX_PATH_LENGTH = 400;
export const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;

/** Returns a user-facing problem with the folder or file name, or null when valid. */
export function validateItemName(name: string, kind: "folder" | "file" = "folder"): string | null {
  const label = kind === "folder" ? "Folder name" : "File name";
  if (!name || !name.trim()) return `${label} is empty.`;
  if (name !== name.trim()) return `${label} can't start or end with a space.`;
  if (name.length > MAX_SEGMENT_LENGTH) return `${label} is longer than ${MAX_SEGMENT_LENGTH} characters.`;
  if ([...name].some((c) => c.charCodeAt(0) < 32 || c.charCodeAt(0) === 127)) return `${label} contains an invisible control character.`;
  if (ILLEGAL_CHARS.test(name)) return `${label} can't contain any of these characters: " * : < > ? / \\ | # %`;
  if (name.endsWith(".")) return `${label} can't end with a period.`;
  if (name.startsWith("~$")) return `${label} can't start with ~$.`;
  if (name.includes("..")) return `${label} can't contain two periods in a row.`;
  if (RESERVED_NAMES.has(name.toLowerCase()) || name.toLowerCase().startsWith("_vti_")) {
    return `"${name}" is a reserved name in SharePoint.`;
  }
  return null;
}

export function buildRootFolderName(input: ProjectInput): string {
  return `${input.projectNumber.trim()} - ${input.projectName.trim()}`;
}

export interface ProjectValidation {
  projectNumber?: string;
  projectName?: string;
  rootFolder?: string;
}

export function validateProjectInput(input: ProjectInput): ProjectValidation {
  const errors: ProjectValidation = {};
  if (!input.projectNumber.trim()) errors.projectNumber = "Enter a project number.";
  else {
    const p = validateItemName(input.projectNumber.trim());
    if (p) errors.projectNumber = p.replace("Folder name", "Project number");
  }
  if (!input.projectName.trim()) errors.projectName = "Enter a project name.";
  else {
    const p = validateItemName(input.projectName.trim());
    if (p) errors.projectName = p.replace("Folder name", "Project name");
  }
  if (!errors.projectNumber && !errors.projectName) {
    const r = validateItemName(buildRootFolderName(input));
    if (r) errors.rootFolder = r;
  }
  return errors;
}

export function hasErrors(v: ProjectValidation): boolean {
  return Boolean(v.projectNumber || v.projectName || v.rootFolder);
}

/** Validates every path of a planned structure, including total path length. */
export function validatePlannedPaths(serverPaths: string[]): string | null {
  const tooLong = serverPaths.find((p) => p.length > MAX_PATH_LENGTH);
  return tooLong
    ? `The full SharePoint path would exceed ${MAX_PATH_LENGTH} characters. Choose a shorter project name or a higher parent folder.`
    : null;
}

/**
 * Normalises a user-entered SharePoint site address to https://host/sites/<name>
 * (or /teams/<name>). Returns null when the address is not a SharePoint site URL.
 * No hostname is assumed: the tenant comes from what the user enters.
 */
export function normalizeSiteUrl(raw: string): string | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;
  let url: URL;
  try {
    url = new URL(/^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`);
  } catch {
    return null;
  }
  if (url.protocol !== "https:") return null;
  if (!/\.sharepoint(-df)?\.[a-z.]+$/i.test(url.hostname)) return null;
  const segments = url.pathname.split("/").filter(Boolean).map((s) => decodeURIComponent(s));
  const managed = segments[0]?.toLowerCase();
  if ((managed === "sites" || managed === "teams") && segments[1]) {
    return `https://${url.hostname.toLowerCase()}/${managed}/${encodeURIComponent(segments[1])}`;
  }
  return null;
}

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function joinPath(...parts: string[]): string {
  return parts
    .map((p) => p.replace(/^\/+|\/+$/g, ""))
    .filter(Boolean)
    .join("/");
}
