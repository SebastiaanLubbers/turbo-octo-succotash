// SharePointFolderService — existence checks and controlled, idempotent folder
// creation. Every requested folder returns exactly one FolderResult.
import type { FolderResult, FolderTemplate, ProjectInput, SharePointConfiguration } from "@/models";
import { AppError, toAppError } from "./errors";
import { GraphItemNotFoundError, GraphSharePoint } from "./graph-sharepoint";
import { buildRootFolderName, joinPath, validateItemName } from "./validation";

export interface PlannedFolder {
  name: string;
  isRoot: boolean;
  /** Relative to the library root (input for CreateNewFolder). */
  relativePath: string;
  /** Server-relative path (input for existence checks and uploads). */
  serverPath: string;
}

export function planStructure(
  config: SharePointConfiguration,
  project: ProjectInput,
  template: FolderTemplate,
): PlannedFolder[] {
  const rootName = buildRootFolderName(project);
  const rootRelative = joinPath(config.parent.relativePath, rootName);
  const rootServer = `/${joinPath(config.parent.serverPath, rootName)}`;
  return [
    { name: rootName, isRoot: true, relativePath: rootRelative, serverPath: rootServer },
    ...template.folders.map((f) => ({
      name: f,
      isRoot: false,
      relativePath: joinPath(rootRelative, f),
      serverPath: `/${joinPath(rootServer, f)}`,
    })),
  ];
}

export type FolderExistence = "exists" | "missing";

export const SharePointFolderService = {
  /** "exists" or "missing"; throws AppError for anything that is not a clear answer. */
  async folderExists(siteUrl: string, serverPath: string): Promise<FolderExistence> {
    try {
      const item = await GraphSharePoint.itemAtPath(siteUrl, serverPath);
      if (!item.folder) {
        throw new AppError("conflict", "A file with this name already exists in that location, so a folder can't be created there.", "GetFolderMetadataByPath returned a file", 409);
      }
      return "exists";
    } catch (error) {
      if (error instanceof GraphItemNotFoundError) return "missing";
      throw toAppError(error, { operation: "GetFolderMetadataByPath", action: "check whether the folder exists" });
    }
  },

  /**
   * Creates the planned structure. Order: root first; subfolders after.
   * - Existing folders are reported "Already exists" and left untouched.
   * - When the root can't be confirmed, its subfolders are "Skipped".
   * - A failed subfolder never stops the remaining subfolders.
   * Running it again only creates what is still missing (idempotent).
   */
  async createStructure(
    config: SharePointConfiguration,
    plan: PlannedFolder[],
    onResult: (index: number, result: FolderResult) => void,
  ): Promise<FolderResult[]> {
    const results: FolderResult[] = [];
    const push = (i: number, r: FolderResult) => {
      results[i] = r;
      onResult(i, r);
    };
    const [root, ...subs] = plan;

    const rootResult = await this.ensureFolder(config, root);
    push(0, rootResult);
    const rootAvailable = rootResult.status === "Created" || rootResult.status === "Already exists";

    for (let i = 0; i < subs.length; i++) {
      const folder = subs[i];
      if (!rootAvailable) {
        push(i + 1, {
          ...base(folder),
          status: "Skipped",
          message: "Skipped because the project folder isn't available.",
        });
        continue;
      }
      push(i + 1, await this.ensureFolder(config, folder));
    }
    return results;
  },

  async ensureFolder(config: SharePointConfiguration, folder: PlannedFolder): Promise<FolderResult> {
    const nameProblem = validateItemName(folder.name);
    if (nameProblem) {
      return { ...base(folder), status: "Failed", message: nameProblem, diagnostic: "Validation failed before any SharePoint call" };
    }
    try {
      const before = await this.folderExists(config.site.url, folder.serverPath);
      if (before === "exists") {
        return { ...base(folder), status: "Already exists", message: "Folder was already there and was kept as is." };
      }
      const location = await GraphSharePoint.resolvePath(config.site.url, folder.serverPath);
      if (location.driveId !== config.library.id || location.relativePath !== folder.relativePath) {
        throw new AppError("validation", "The saved SharePoint location has changed. Choose the location again.", "The selected library does not match the requested folder path");
      }
      try {
        await GraphSharePoint.createFolder(config.library.id, folder.relativePath);
      } catch (error) {
        const err = toAppError(error, { operation: "CreateNewFolder", action: "create this folder" });
        // A parallel run may have created it in the meantime: re-check before reporting failure.
        if (err.kind === "conflict" && (await this.folderExists(config.site.url, folder.serverPath)) === "exists") {
          return { ...base(folder), status: "Already exists", message: "Folder was created by someone else in the meantime and was kept." };
        }
        return { ...base(folder), status: "Failed", message: err.userMessage, diagnostic: err.diagnostic };
      }
      // Verify — success is reported only when SharePoint shows the folder at the exact path.
      const after = await this.folderExists(config.site.url, folder.serverPath);
      if (after !== "exists") {
        return {
          ...base(folder),
          status: "Failed",
          message: "SharePoint accepted the request but the folder isn't at the expected location. Check the library and try again.",
          diagnostic: "CreateNewFolder succeeded; GetFolderMetadataByPath returned 404",
        };
      }
      return { ...base(folder), status: "Created", message: "Folder created." };
    } catch (e) {
      const err = toAppError(e, { operation: "CreateNewFolder", action: "create this folder" });
      return { ...base(folder), status: "Failed", message: err.userMessage, diagnostic: err.diagnostic };
    }
  },
};

function base(folder: PlannedFolder) {
  return { name: folder.name, relativePath: folder.relativePath, serverPath: folder.serverPath, isRoot: folder.isRoot };
}
