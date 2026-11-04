import { useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { ChevronRight, ExternalLink, FileText, Folder, FolderOpen, FolderTree, RotateCw, Search, Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import { EmptyState } from "@/components/states";
import { AddFilePanel } from "@/components/explorer/add-file-panel";
import { ErrorPanel } from "@/components/wizard/error-panel";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";
import { toAppError } from "@/services/errors";
import { SharePointSiteService, type FolderEntry } from "@/services/sharepoint-site-service";
import { formatBytes } from "@/services/validation";
import { sharePointLink, splitRel } from "@/state/selectors";
import { useWizard } from "@/state/wizard-store";

function useContents(siteUrl: string | undefined, folderId: string | undefined) {
  return useQuery({
    queryKey: ["contents", siteUrl, folderId],
    queryFn: async () => {
      try {
        return await SharePointSiteService.listContents(siteUrl!, folderId!);
      } catch (e) {
        throw toAppError(e, { operation: "ListFolder", action: "read the folder contents" });
      }
    },
    enabled: Boolean(siteUrl && folderId),
    retry: false,
  });
}

function formatDate(value?: string): string {
  if (!value) return "";
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? "" : d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
}

const linkClass =
  "inline-flex items-center gap-1.5 text-sm font-medium text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring";

/** Existing projects: browse folders and files below the configured parent and add files to them. */
export function ProjectsPage() {
  const config = useWizard((s) => s.config);
  const goTo = useWizard((s) => s.goTo);
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const [filter, setFilter] = useState("");
  const [trail, setTrail] = useState<FolderEntry[]>([]);
  const [adding, setAdding] = useState(false);

  const projects = useContents(config?.site.url, config?.parent.folderId);
  const current = trail[trail.length - 1];
  const contents = useContents(config?.site.url, current?.id);

  const visibleProjects = useMemo(() => {
    const q = filter.trim().toLowerCase();
    // The upload-flow queue folder (library root) is infrastructure, not a project.
    const queue = config?.uploadFlow?.queueFolder?.toLowerCase();
    const list = (projects.data?.folders ?? []).filter((f) => !(queue && !config?.parent.relativePath && f.name.toLowerCase() === queue));
    return q ? list.filter((p) => p.name.toLowerCase().includes(q)) : list;
  }, [projects.data, filter, config]);

  if (!config) {
    return (
      <div className="mx-auto w-full max-w-3xl px-4 py-10">
        <EmptyState
          icon={FolderTree}
          title="No SharePoint location configured"
          description="Set up the site, library and parent folder first."
          action={<Button onClick={() => { goTo(1); navigate("/"); }}>Start setup</Button>}
        />
      </div>
    );
  }

  const baseParts = [config.library.displayName, ...splitRel(config.parent.relativePath)];
  const pathParts = [...baseParts, ...trail.map((t) => t.name)];
  const currentServerPath = current?.serverPath || "";

  function refresh() {
    void queryClient.invalidateQueries({ queryKey: ["contents", config!.site.url] });
  }
  function openProject(p: FolderEntry) {
    setTrail([p]);
    setAdding(false);
  }

  return (
    <div className="mx-auto grid w-full max-w-[1400px] gap-6 px-4 py-6 md:grid-cols-[18rem_minmax(0,1fr)] md:px-6 lg:py-8">
      {/* Project list */}
      <aside aria-labelledby="projects-title" className="min-w-0 space-y-3 md:sticky md:top-20 md:self-start">
        <div className="flex items-center justify-between gap-2">
          <h1 id="projects-title" className="text-xl font-semibold">Existing projects</h1>
          <Button variant="ghost" size="icon" onClick={refresh} aria-label="Refresh" disabled={projects.isFetching}>
            <RotateCw className={cn(projects.isFetching && "animate-spin")} aria-hidden="true" />
          </Button>
        </div>
        <p className="text-xs break-words text-muted-foreground">{[config.site.label, ...baseParts].join(" › ")}</p>
        <div className="relative">
          <Search className="pointer-events-none absolute top-1/2 left-2.5 size-4 -translate-y-1/2 text-muted-foreground" aria-hidden="true" />
          <Input value={filter} onChange={(e) => setFilter(e.target.value)} placeholder="Filter projects" aria-label="Filter projects" className="pl-8" />
        </div>
        {projects.isPending && <div role="status" aria-label="Loading projects" className="space-y-2">{[0, 1, 2, 3].map((i) => <Skeleton key={i} className="h-9 w-full" />)}</div>}
        {projects.isError && <ErrorPanel title="Projects couldn't be loaded" error={projects.error} action={<Button size="sm" variant="outline" onClick={() => projects.refetch()}>Try again</Button>} />}
        {projects.isSuccess && (projects.data.folders.length === 0 ? (
          <p className="rounded-lg border p-3 text-sm">No project folders yet.</p>
        ) : visibleProjects.length === 0 ? (
          <p className="rounded-lg border p-3 text-sm">No project matches "{filter}".</p>
        ) : (
          <nav aria-label="Projects">
            <ul className="max-h-[65vh] space-y-0.5 overflow-auto rounded-lg border p-1">
              {visibleProjects.map((p) => {
                const selected = trail[0]?.id === p.id;
                return (
                  <li key={p.id}>
                    <button
                      type="button"
                      onClick={() => openProject(p)}
                      aria-current={selected ? "true" : undefined}
                      className={cn(
                        "flex w-full items-center gap-2 rounded-md px-2 py-2 text-left text-sm hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring",
                        selected && "bg-accent font-medium",
                      )}
                    >
                      {selected ? <FolderOpen className="size-4 shrink-0 text-primary" aria-hidden="true" /> : <Folder className="size-4 shrink-0 text-muted-foreground" aria-hidden="true" />}
                      <span className="min-w-0 truncate" title={p.name}>{p.name}</span>
                    </button>
                  </li>
                );
              })}
            </ul>
          </nav>
        ))}
        {projects.isSuccess && <p className="text-xs text-muted-foreground">{projects.data.folders.length} folders</p>}
      </aside>

      {/* Folder contents */}
      <section id="wizard-content" tabIndex={-1} aria-labelledby="folder-title" className="min-w-0 space-y-5 outline-none">
        {!current ? (
          <EmptyState icon={FolderOpen} title="Select a project" description="Its folders and files appear here." />
        ) : (
          <>
            <div className="space-y-3">
              <nav aria-label="Folder path" className="flex flex-wrap items-center gap-1 text-sm">
                {trail.map((t, i) => (
                  <span key={t.id} className="flex items-center gap-1">
                    {i > 0 && <ChevronRight className="size-3.5 text-muted-foreground" aria-hidden="true" />}
                    <button
                      type="button"
                      onClick={() => { setTrail(trail.slice(0, i + 1)); setAdding(false); }}
                      aria-current={i === trail.length - 1 ? "location" : undefined}
                      className={cn("rounded px-1 hover:underline focus-visible:outline-2 focus-visible:outline-ring", i === trail.length - 1 && "font-medium")}
                    >
                      {t.name}
                    </button>
                  </span>
                ))}
              </nav>
              <div className="flex flex-wrap items-center gap-3">
                <h2 id="folder-title" className="min-w-0 flex-1 truncate text-2xl font-semibold tracking-tight">{current.name}</h2>
                {currentServerPath && (
                  <a href={safeUrl(sharePointLink(config.site.url, currentServerPath)) ?? "#"} target="_blank" rel="noreferrer" className={linkClass}>
                    Open in SharePoint <ExternalLink className="size-3.5" aria-hidden="true" />
                    <span className="sr-only">(opens in a new tab)</span>
                  </a>
                )}
                <Button onClick={() => setAdding(true)} disabled={adding || !currentServerPath}>
                  <Upload aria-hidden="true" /> Add file
                </Button>
              </div>
            </div>

            {adding && currentServerPath && (
              <AddFilePanel
                key={current.id}
                config={config}
                folderServerPath={currentServerPath}
                pathParts={[config.site.label, ...pathParts]}
                onClose={() => setAdding(false)}
                onUploaded={refresh}
              />
            )}

            {contents.isPending && <div role="status" aria-label="Loading folder" className="space-y-2">{[0, 1, 2, 3].map((i) => <Skeleton key={i} className="h-11 w-full" />)}</div>}
            {contents.isError && <ErrorPanel title="Folder couldn't be loaded" error={contents.error} action={<Button size="sm" variant="outline" onClick={() => contents.refetch()}>Try again</Button>} />}
            {contents.isSuccess && (
              <div className="overflow-hidden rounded-lg border">
                <table className="w-full text-sm">
                  <caption className="sr-only">Contents of {current.name}</caption>
                  <thead className="bg-muted text-left">
                    <tr>
                      <th scope="col" className="px-3 py-2 font-medium">Name</th>
                      <th scope="col" className="hidden px-3 py-2 font-medium sm:table-cell">Modified</th>
                      <th scope="col" className="px-3 py-2 text-right font-medium">Size</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {contents.data.folders.map((f) => (
                      <tr key={f.id} className="hover:bg-muted/50">
                        <td className="px-3 py-2">
                          <button
                            type="button"
                            onClick={() => { setTrail([...trail, f]); setAdding(false); }}
                            className="flex min-w-0 items-center gap-2 rounded text-left hover:underline focus-visible:outline-2 focus-visible:outline-ring"
                          >
                            <Folder className="size-4 shrink-0 text-primary" aria-hidden="true" />
                            <span className="truncate">{f.name}</span>
                            <span className="sr-only">(folder)</span>
                          </button>
                        </td>
                        <td className="hidden px-3 py-2 sm:table-cell" />
                        <td className="px-3 py-2 text-right">—</td>
                      </tr>
                    ))}
                    {contents.data.files.map((f) => (
                      <tr key={f.id} className="hover:bg-muted/50">
                        <td className="px-3 py-2">
                          <a
                            href={safeUrl(sharePointLink(config.site.url, f.serverPath)) ?? "#"}
                            target="_blank"
                            rel="noreferrer"
                            className="flex min-w-0 items-center gap-2 rounded hover:underline focus-visible:outline-2 focus-visible:outline-ring"
                          >
                            <FileText className="size-4 shrink-0 text-muted-foreground" aria-hidden="true" />
                            <span className="truncate">{f.name}</span>
                            <span className="sr-only">(file, opens in a new tab)</span>
                          </a>
                        </td>
                        <td className="hidden px-3 py-2 whitespace-nowrap sm:table-cell">{formatDate(f.lastModified)}</td>
                        <td className="px-3 py-2 text-right whitespace-nowrap tabular-nums">{typeof f.size === "number" ? formatBytes(f.size) : ""}</td>
                      </tr>
                    ))}
                    {contents.data.folders.length + contents.data.files.length === 0 && (
                      <tr>
                        <td colSpan={3} className="px-3 py-6 text-center">This folder is empty.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            )}
            {contents.isSuccess && (
              <p className="text-xs text-muted-foreground">{contents.data.folders.length} folders · {contents.data.files.length} files</p>
            )}
          </>
        )}
      </section>
    </div>
  );
}
