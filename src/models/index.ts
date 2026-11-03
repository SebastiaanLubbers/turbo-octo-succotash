// Domain models for SharePoint Structure Builder.
// These types carry NON-SECRET metadata only: no tokens, passwords or user names.

/** A SharePoint site the signed-in user can reach. */
export interface SharePointSite {
  /** Absolute site URL, e.g. https://<tenant>.sharepoint.com/sites/<site>. Entered by the user at runtime. */
  url: string;
  /** Display label derived from the URL (last path segment). */
  label: string;
}

/** A document library on the selected site. */
export interface DocumentLibrary {
  /** Library identifier used by folder operations (list GUID when available, otherwise display name). */
  id: string;
  displayName: string;
  /** Opaque folder id of the library root, used to browse folders. */
  rootFolderId: string;
  /** Server-relative path of the library root, e.g. "/Shared Documents". */
  rootPath: string;
}

/** The library root or an existing folder below it. */
export interface ParentLocation {
  /** Opaque folder id used to browse its children. */
  folderId: string;
  /** Path relative to the library root ("" = library root). */
  relativePath: string;
  /** Server-relative path, e.g. "/Shared Documents/Projects". */
  serverPath: string;
}

/** Saved, non-secret configuration. Stored per user in the browser; never in source code. */
export interface SharePointConfiguration {
  version: 1;
  site: SharePointSite;
  library: DocumentLibrary;
  parent: ParentLocation;
  savedAt: string;
}

export interface ProjectInput {
  projectNumber: string;
  projectName: string;
}

export interface FolderTemplate {
  id: string;
  name: string;
  description: string;
  /** Ordered subfolder names created under the project root. */
  folders: string[];
}

export type FolderResultStatus = "Created" | "Already exists" | "Failed" | "Skipped";

export interface FolderResult {
  /** Folder name as requested. */
  name: string;
  /** Path relative to the library root. */
  relativePath: string;
  /** Server-relative path. */
  serverPath: string;
  isRoot: boolean;
  status: FolderResultStatus;
  /** Business-friendly explanation shown to the user. */
  message: string;
  /** Developer diagnostic (operation, HTTP status). Never contains tokens or raw payloads. */
  diagnostic?: string;
}

export interface UploadInput {
  file: File;
  /** Destination folder (server-relative path). */
  folderServerPath: string;
  /** Destination folder relative to the library root, for display. */
  folderRelativePath: string;
  /** Final file name in SharePoint (may differ from file.name when the user chose a new name). */
  targetFileName: string;
}

export type UploadStatus = "Uploaded" | "Failed" | "Blocked";

export interface UploadResult {
  status: UploadStatus;
  fileName: string;
  serverPath: string;
  sizeBytes: number;
  message: string;
  diagnostic?: string;
  /** True when the upload was blocked because the file name already exists. */
  duplicate?: boolean;
  /** True when the upload was blocked because the file isn't a text file. */
  binary?: boolean;
}
