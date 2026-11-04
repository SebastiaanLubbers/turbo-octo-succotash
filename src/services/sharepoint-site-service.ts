// SharePointSiteService — site access check, document library discovery and
// folder browsing, through Microsoft Graph using the
// signed-in user's own permissions.
import type { DocumentLibrary, SharePointSite } from "@/models";
import { AppError } from "./errors";
import { GraphSharePoint, itemKey, siteRelativePath } from "./graph-sharepoint";
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
    await GraphSharePoint.siteId(url);
    return { url, label: siteLabel(url) };
  },

  /** Document libraries of a site: root folders that match a list/library of the site. */
  async listLibraries(siteUrl: string): Promise<DocumentLibrary[]> {
    return (await GraphSharePoint.libraries(siteUrl)).sort((a, b) => a.displayName.localeCompare(b.displayName));
  },

  /** Subfolders and files of a folder, each sorted by name. */
  async listContents(siteUrl: string, folderId: string): Promise<FolderContents> {
    const { driveId, items } = await GraphSharePoint.children(folderId);
    const byName = (a: { name: string }, b: { name: string }) => a.name.localeCompare(b.name, undefined, { numeric: true });
    const folders: FolderEntry[] = [];
    const files: FileEntry[] = [];
    for (const i of items) {
      if (!i.id || !i.name || !i.webUrl) continue;
      const entry = { id: itemKey(driveId, i.id), name: i.name, serverPath: siteRelativePath(siteUrl, i.webUrl) };
      if (i.folder) {
        if (i.name === "Forms") continue;
        folders.push(entry);
      } else if (i.file) {
        files.push({ ...entry, size: i.size, lastModified: i.lastModifiedDateTime, mediaType: i.file.mimeType });
      }
    }
    return { folders: folders.sort(byName), files: files.sort(byName) };
  },

  /** Subfolders of a folder (library root or any folder below it). */
  async listFolders(siteUrl: string, folderId: string): Promise<FolderEntry[]> {
    return (await this.listContents(siteUrl, folderId)).folders;
  },
};
