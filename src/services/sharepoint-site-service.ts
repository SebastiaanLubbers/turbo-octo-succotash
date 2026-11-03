// SharePointSiteService — site access check, document library discovery and
// folder browsing, through the app's bound SharePoint connector using the
// signed-in user's own permissions.
import { SharePointService } from "../../generated/services/SharePointService";
import type { BlobMetadata, Table } from "../../generated/models/SharePointModel";
import type { DocumentLibrary, SharePointSite } from "@/models";
import { AppError, unwrap } from "./errors";
import { normalizeSiteUrl } from "./validation";

export interface FileEntry {
  id: string;
  name: string;
  serverPath: string;
  size?: number;
  lastModified?: string;
  mediaType?: string;
}

export interface FolderContents {
  folders: FolderEntry[];
  files: FileEntry[];
}

export interface FolderEntry {
  id: string;
  name: string;
  serverPath: string;
}

function siteLabel(url: string): string {
  const last = url.split("/").filter(Boolean).pop() ?? url;
  try {
    return decodeURIComponent(last);
  } catch {
    return last;
  }
}

function trimPath(p: string | undefined): string {
  return (p ?? "").replace(/\/+$/, "");
}

export const SharePointSiteService = {
  /** Checks that the site exists and the user can read it. */
  async openSite(rawUrl: string): Promise<SharePointSite> {
    const url = normalizeSiteUrl(rawUrl);
    if (!url) {
      throw new AppError(
        "validation",
        "Enter a SharePoint site address such as https://<your-tenant>.sharepoint.com/sites/<site-name>.",
        "normalizeSiteUrl rejected the input",
      );
    }
    const result = await SharePointService.GetAllTables(url);
    unwrap(result, { operation: "GetAllTables", action: "open this site" });
    return { url, label: siteLabel(url) };
  },

  /** Document libraries of a site: root folders that match a list/library of the site. */
  async listLibraries(siteUrl: string): Promise<DocumentLibrary[]> {
    const [tablesResult, rootResult] = await Promise.all([
      SharePointService.GetAllTables(siteUrl),
      SharePointService.ListRootFolder(siteUrl),
    ]);
    const tables: Table[] = unwrap(tablesResult, { operation: "GetAllTables", action: "read the document libraries" }).value ?? [];
    const root: BlobMetadata[] = unwrap(rootResult, { operation: "ListRootFolder", action: "read the document libraries" }) ?? [];
    const byName = new Map<string, Table>();
    for (const t of tables) {
      if (t.DisplayName) byName.set(t.DisplayName.toLowerCase(), t);
    }
    const libraries: DocumentLibrary[] = [];
    for (const f of root) {
      if (!f.IsFolder || !f.Id) continue;
      const display = f.DisplayName || f.Name || "";
      const table = byName.get(display.toLowerCase()) ?? (f.Name ? byName.get(f.Name.toLowerCase()) : undefined);
      if (!table) continue; // a system folder, not a library
      libraries.push({
        id: table.Name || display,
        displayName: table.DisplayName || display,
        rootFolderId: f.Id,
        rootPath: trimPath(f.Path) || `/${f.Name ?? display}`,
      });
    }
    return libraries.sort((a, b) => a.displayName.localeCompare(b.displayName));
  },

  /** Subfolders and files of a folder, each sorted by name. */
  async listContents(siteUrl: string, folderId: string): Promise<FolderContents> {
    const result = await SharePointService.ListFolder(siteUrl, folderId);
    const items = unwrap(result, { operation: "ListFolder", action: "read the folder contents" }) ?? [];
    const byName = (a: { name: string }, b: { name: string }) => a.name.localeCompare(b.name, undefined, { numeric: true });
    const folders: FolderEntry[] = [];
    const files: FileEntry[] = [];
    for (const i of items) {
      const name = i.DisplayName || i.Name;
      if (!i.Id || !name) continue;
      if (i.IsFolder) {
        if ((i.Name ?? "") === "Forms") continue;
        folders.push({ id: i.Id, name, serverPath: trimPath(i.Path) });
      } else {
        files.push({ id: i.Id, name, serverPath: trimPath(i.Path), size: i.Size, lastModified: i.LastModified, mediaType: i.MediaType });
      }
    }
    return { folders: folders.sort(byName), files: files.sort(byName) };
  },

  /** Subfolders of a folder (library root or any folder below it). */
  async listFolders(siteUrl: string, folderId: string): Promise<FolderEntry[]> {
    const result = await SharePointService.ListFolder(siteUrl, folderId);
    const items = unwrap(result, { operation: "ListFolder", action: "read the folders" }) ?? [];
    return items
      .filter((i) => i.IsFolder && i.Id && (i.Name || i.DisplayName) && (i.Name ?? "") !== "Forms")
      .map((i) => ({
        id: i.Id as string,
        name: (i.DisplayName || i.Name) as string,
        serverPath: trimPath(i.Path),
      }))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  },
};
