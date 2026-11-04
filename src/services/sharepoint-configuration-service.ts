// SharePointConfigurationService — persists the NON-SECRET configuration
// (site URL, library, parent folder) chosen in the setup wizard.
// Storage: the browser's localStorage, per user and per app origin. Nothing
// customer-specific is compiled into the source; another customer simply runs
// the wizard against their own tenant.
import type { SharePointConfiguration } from "@/models";

const STORAGE_KEY = "spsb.graph.configuration.v1";
const RECENT_SITES_KEY = "spsb.recentSites.v1";

function isConfiguration(value: unknown): value is SharePointConfiguration {
  if (!value || typeof value !== "object") return false;
  const v = value as Partial<SharePointConfiguration>;
  return (
    v.version === 1 &&
    typeof v.site?.url === "string" &&
    typeof v.library?.id === "string" &&
    typeof v.library?.rootPath === "string" &&
    typeof v.parent?.serverPath === "string" &&
    typeof v.parent?.relativePath === "string"
  );
}

export const SharePointConfigurationService = {
  load(): SharePointConfiguration | null {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      const parsed: unknown = JSON.parse(raw);
      return isConfiguration(parsed) ? parsed : null;
    } catch {
      return null;
    }
  },

  save(config: Omit<SharePointConfiguration, "version" | "savedAt">): SharePointConfiguration {
    const full: SharePointConfiguration = { ...config, version: 1, savedAt: new Date().toISOString() };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(full));
    this.rememberSite(config.site.url);
    return full;
  },

  clear(): void {
    localStorage.removeItem(STORAGE_KEY);
  },

  recentSites(): string[] {
    try {
      const raw = localStorage.getItem(RECENT_SITES_KEY);
      const parsed: unknown = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed.filter((x): x is string => typeof x === "string").slice(0, 8) : [];
    } catch {
      return [];
    }
  },

  rememberSite(url: string): void {
    const next = [url, ...this.recentSites().filter((u) => u !== url)].slice(0, 8);
    localStorage.setItem(RECENT_SITES_KEY, JSON.stringify(next));
  },
};
