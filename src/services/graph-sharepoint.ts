import { graphAccessToken } from "./browser-auth";

const graphRoot = "https://graph.microsoft.com/v1.0";

export interface GraphItem {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  lastModifiedDateTime?: string;
  folder?: object;
  file?: { mimeType?: string };
}

interface GraphDrive {
  id: string;
  name: string;
  driveType?: string;
}

interface GraphList<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export class GraphHttpError extends Error {
  constructor(readonly status: number) {
    super(`Microsoft Graph request failed (HTTP ${status})`);
  }
}

export class GraphItemNotFoundError extends GraphHttpError {
  constructor() { super(404); }
}

async function graphRequest<T>(path: string, init?: RequestInit): Promise<T> {
  const url = new URL(path, graphRoot + "/");
  if (url.origin !== "https://graph.microsoft.com" || !url.pathname.startsWith("/v1.0/")) {
    throw new Error("Invalid Microsoft Graph URL.");
  }
  const token = await graphAccessToken();
  const response = await fetch(url, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, ...init?.headers },
  });
  if (!response.ok) throw new GraphHttpError(response.status);
  return response.status === 204 ? undefined as T : await response.json() as T;
}

async function graphList<T>(path: string): Promise<T[]> {
  const items: T[] = [];
  let next: string | undefined = path;
  while (next) {
    const page: GraphList<T> = await graphRequest<GraphList<T>>(next);
    items.push(...page.value);
    next = page["@odata.nextLink"];
  }
  return items;
}

function encodePath(path: string): string {
  return path.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

export function itemKey(driveId: string, itemId: string): string {
  return `${driveId}|${itemId}`;
}

function splitItemKey(key: string): [string, string] {
  const parts = key.split("|");
  if (parts.length !== 2 || !parts[0] || !parts[1]) throw new Error("Select the SharePoint location again.");
  return [parts[0], parts[1]];
}

export function siteRelativePath(siteUrl: string, webUrl: string): string {
  const site = new URL(siteUrl);
  const item = new URL(webUrl);
  if (site.origin !== item.origin) throw new Error("SharePoint returned an unexpected item URL.");
  const sitePath = decodeURIComponent(site.pathname).replace(/\/+$/, "");
  const itemPath = decodeURIComponent(item.pathname);
  if (!itemPath.toLowerCase().startsWith(sitePath.toLowerCase() + "/")) throw new Error("SharePoint returned an item outside the selected site.");
  return itemPath.slice(sitePath.length).replace(/\/+$/, "");
}

export const GraphSharePoint = {
  async siteId(siteUrl: string): Promise<string> {
    const site = new URL(siteUrl);
    const result = await graphRequest<{ id: string }>(`sites/${site.hostname}:${site.pathname}`);
    return result.id;
  },

  async libraries(siteUrl: string) {
    const siteId = await this.siteId(siteUrl);
    const drives = await graphList<GraphDrive>(`sites/${encodeURIComponent(siteId)}/drives`);
    return Promise.all(drives.filter((drive) => drive.driveType === "documentLibrary").map(async (drive) => {
      const root = await graphRequest<GraphItem>(`drives/${encodeURIComponent(drive.id)}/root`);
      return { id: drive.id, displayName: drive.name, rootFolderId: itemKey(drive.id, root.id), rootPath: siteRelativePath(siteUrl, root.webUrl) };
    }));
  },

  async children(folderId: string): Promise<{ driveId: string; items: GraphItem[] }> {
    const [driveId, itemId] = splitItemKey(folderId);
    const items = await graphList<GraphItem>(`drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children`);
    return { driveId, items };
  },

  async resolvePath(siteUrl: string, serverPath: string): Promise<{ driveId: string; relativePath: string }> {
    const path = "/" + serverPath.replace(/^\/+|\/+$/g, "");
    const libraries = await this.libraries(siteUrl);
    const library = libraries.sort((a, b) => b.rootPath.length - a.rootPath.length).find((candidate) =>
      path.toLowerCase() === candidate.rootPath.toLowerCase() || path.toLowerCase().startsWith(candidate.rootPath.toLowerCase() + "/"));
    if (!library) throw new GraphHttpError(404);
    return { driveId: library.id, relativePath: path.slice(library.rootPath.length).replace(/^\//, "") };
  },

  async itemAtPath(siteUrl: string, serverPath: string): Promise<GraphItem> {
    const { driveId, relativePath } = await this.resolvePath(siteUrl, serverPath);
    const endpoint = relativePath ? `root:/${encodePath(relativePath)}` : "root";
    try {
      return await graphRequest<GraphItem>(`drives/${encodeURIComponent(driveId)}/${endpoint}`);
    } catch (error) {
      if (error instanceof GraphHttpError && error.status === 404) throw new GraphItemNotFoundError();
      throw error;
    }
  },

  async createFolder(driveId: string, relativePath: string): Promise<void> {
    const parts = relativePath.split("/").filter(Boolean);
    const name = parts.pop();
    if (!name) throw new Error("Folder name is missing.");
    const parent = parts.length ? `root:/${encodePath(parts.join("/"))}:/children` : "root/children";
    await graphRequest(`drives/${encodeURIComponent(driveId)}/${parent}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    });
  },

  async upload(siteUrl: string, serverPath: string, file: File): Promise<GraphItem> {
    const { driveId, relativePath } = await this.resolvePath(siteUrl, serverPath);
    const session = await graphRequest<{ uploadUrl: string }>(`drives/${encodeURIComponent(driveId)}/root:/${encodePath(relativePath)}:/createUploadSession`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "fail" } }),
    });
    const uploadUrl = new URL(session.uploadUrl);
    if (uploadUrl.protocol !== "https:") throw new Error("Invalid upload session URL.");
    const response = await fetch(uploadUrl, {
      method: "PUT",
      headers: { "Content-Range": `bytes 0-${file.size - 1}/${file.size}` },
      body: file,
    });
    if (!response.ok) throw new GraphHttpError(response.status);
    if (response.status === 202) throw new Error("The upload was not completed.");
    return await response.json() as GraphItem;
  },
};