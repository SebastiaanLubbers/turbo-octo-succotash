import { useState } from "react";
import { Link } from "react-router-dom";
import { useQuery } from "@tanstack/react-query";
import { ArrowRight, ChevronRight, FolderOpen, Globe, Library, Settings2, UserRound } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Skeleton } from "@/components/ui/skeleton";
import { Spinner } from "@/components/ui/spinner";
import { EmptyState } from "@/components/states";
import { SharePointConfigurationService } from "@/services/sharepoint-configuration-service";
import { DEFAULT_QUEUE_FOLDER } from "@/services/upload-flow-service";
import { validateItemName } from "@/services/validation";
import { Checkbox } from "@/components/ui/checkbox";
import { SharePointSiteService, type FolderEntry } from "@/services/sharepoint-site-service";
import { useBrowserAuth } from "@/services/browser-auth";
import { toAppError } from "@/services/errors";
import { useWizard } from "@/state/wizard-store";
import { splitRel } from "@/state/selectors";
import { ErrorPanel } from "./error-panel";
import { PathLine } from "./destination-tree";
import { Field, StepFrame } from "./step-frame";

// ---------- 1. Welcome ----------
export function WelcomeStep() {
  const config = useWizard((s) => s.config);
  const goTo = useWizard((s) => s.goTo);
  const startReconfigure = useWizard((s) => s.startReconfigure);
  const startNewProject = useWizard((s) => s.startNewProject);
  return (
    <StepFrame
      title="SharePoint Structure Builder"
      primary={
        config ? (
          <Button onClick={startNewProject}>
            Start new project <ArrowRight aria-hidden="true" />
          </Button>
        ) : (
          <Button onClick={() => goTo(2)}>
            Start setup <ArrowRight aria-hidden="true" />
          </Button>
        )
      }
      secondary={
        config ? (
          <>
            <Button variant="outline" onClick={startReconfigure}>
              <Settings2 aria-hidden="true" /> Change configuration
            </Button>
            <Button variant="outline" asChild>
              <Link to="/projects">
                <FolderOpen aria-hidden="true" /> Existing projects
              </Link>
            </Button>
          </>
        ) : undefined
      }
    >
      <ol className="grid gap-3 sm:grid-cols-3">
        {[
          { n: "1", t: "Choose a location", d: "Site, document library and parent folder" },
          { n: "2", t: "Create the structure", d: "Project folder with the template's subfolders" },
          { n: "3", t: "Upload a document", d: "Into one of the project folders" },
        ].map((s) => (
          <li key={s.n} className="rounded-lg border bg-card p-4">
            <span className="flex size-7 items-center justify-center rounded-full bg-primary text-sm font-semibold text-primary-foreground" aria-hidden="true">
              {s.n}
            </span>
            <p className="mt-3 font-medium">{s.t}</p>
            <p className="mt-1 text-sm text-muted-foreground">{s.d}</p>
          </li>
        ))}
      </ol>
      {config ? (
        <div className="rounded-lg border p-4">
          <p className="text-sm font-semibold">Saved location</p>
          <div className="mt-2">
            <PathLine parts={[config.site.label, config.library.displayName, ...splitRel(config.parent.relativePath)]} />
          </div>
          <p className="mt-1 text-xs break-all text-muted-foreground">{config.site.url}</p>
        </div>
      ) : (
        <p className="text-sm">No SharePoint location is configured yet.</p>
      )}
    </StepFrame>
  );
}

// ---------- 2. Microsoft 365 connection ----------
export function ConnectionStep() {
  const goTo = useWizard((s) => s.goTo);
  const account = useBrowserAuth((state) => state.account);
  return (
    <StepFrame title="Microsoft 365 connection" onBack={() => goTo(1)} primary={<Button onClick={() => goTo(3)} disabled={!account}>Continue <ArrowRight aria-hidden="true" /></Button>}>
      {account && (
        <div className="flex items-center gap-4 rounded-lg border bg-card p-4">
          <span className="flex size-10 items-center justify-center rounded-full bg-accent" aria-hidden="true">
            <UserRound className="size-5" />
          </span>
          <div className="min-w-0">
            <p className="text-sm text-muted-foreground">Working as</p>
            <p className="truncate font-medium">{account.name || account.username}</p>
          </div>
        </div>
      )}
      <ul className="space-y-2 text-sm">
        <li>SharePoint is accessed with your own Microsoft 365 account and permissions.</li>
        <li>Microsoft handles sign-in; your browser keeps the session until you sign out or close the tab.</li>
        <li>Only the chosen site address, library and folder are saved, in this browser.</li>
      </ul>
    </StepFrame>
  );
}

// ---------- 3. Select SharePoint site ----------
export function SiteStep() {
  const site = useWizard((s) => s.site);
  const setSite = useWizard((s) => s.setSite);
  const goTo = useWizard((s) => s.goTo);
  const [query, setQuery] = useState(site?.url ?? "");
  const [opening, setOpening] = useState<string | null>(null);
  const [error, setError] = useState<unknown>(null);
  const recent = SharePointConfigurationService.recentSites();
  const filtered = query.trim() ? recent.filter((u) => u.toLowerCase().includes(query.trim().toLowerCase())) : recent;

  async function open(address: string) {
    setOpening(address);
    setError(null);
    try {
      const s = await SharePointSiteService.openSite(address);
      setSite(s);
      SharePointConfigurationService.rememberSite(s.url);
      goTo(4);
    } catch (e) {
      setError(toAppError(e, { operation: "GetAllTables", action: "open this site" }));
    } finally {
      setOpening(null);
    }
  }

  return (
    <StepFrame title="Select SharePoint site" onBack={() => goTo(2)} primary={site ? <Button variant="outline" onClick={() => goTo(4)}>Keep {site.label} <ArrowRight aria-hidden="true" /></Button> : undefined}>
      <form
        className="space-y-3"
        onSubmit={(e) => {
          e.preventDefault();
          void open(query);
        }}
      >
        <Field label="Site address" htmlFor="site-url" hint="Example: https://<tenant>.sharepoint.com/sites/<site-name>">
          <div className="flex gap-2">
            <Input
              id="site-url"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="https://…sharepoint.com/sites/…"
              autoComplete="url"
              aria-describedby="site-url-hint"
            />
            <Button type="submit" disabled={!query.trim() || opening !== null}>
              {opening === query ? <Spinner /> : <Globe aria-hidden="true" />} Open site
            </Button>
          </div>
        </Field>
      </form>
      {error !== null && <ErrorPanel title="Site not accessible" error={error} />}
      {filtered.length > 0 && (
        <div>
          <h2 className="mb-2 text-sm font-semibold">Recent sites</h2>
          <ul className="divide-y rounded-lg border">
            {filtered.map((u) => (
              <li key={u}>
                <button
                  type="button"
                  onClick={() => void open(u)}
                  disabled={opening !== null}
                  className="flex w-full items-center gap-3 px-3 py-2.5 text-left text-sm hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring"
                >
                  <Globe className="size-4 shrink-0 text-muted-foreground" aria-hidden="true" />
                  <span className="min-w-0 flex-1 truncate">{u}</span>
                  {opening === u ? <Spinner /> : <ChevronRight className="size-4 text-muted-foreground" aria-hidden="true" />}
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </StepFrame>
  );
}

// ---------- 4. Select document library ----------
export function LibraryStep() {
  const site = useWizard((s) => s.site);
  const library = useWizard((s) => s.library);
  const setLibrary = useWizard((s) => s.setLibrary);
  const goTo = useWizard((s) => s.goTo);
  const libs = useQuery({
    queryKey: ["libraries", site?.url],
    queryFn: async () => {
      try {
        return await SharePointSiteService.listLibraries(site!.url);
      } catch (e) {
        throw toAppError(e, { operation: "GetAllTables/ListRootFolder", action: "read the document libraries" });
      }
    },
    enabled: Boolean(site),
    retry: false,
  });
  return (
    <StepFrame
      title="Select document library"
      onBack={() => goTo(3)}
      primary={<Button disabled={!library} onClick={() => goTo(5)}>Continue <ArrowRight aria-hidden="true" /></Button>}
    >
      {site && <PathLine parts={[site.label]} />}
      {libs.isPending && <div role="status" aria-label="Loading libraries" className="space-y-2">{[0, 1, 2].map((i) => <Skeleton key={i} className="h-12 w-full" />)}</div>}
      {libs.isError && <ErrorPanel title="Libraries couldn't be loaded" error={libs.error} action={<Button size="sm" variant="outline" onClick={() => libs.refetch()}>Try again</Button>} />}
      {libs.isSuccess && libs.data.length === 0 && (
        <EmptyState icon={Library} title="No document library available" description="This site has no document library you can access. Choose another site or ask the site owner for access." />
      )}
      {libs.isSuccess && libs.data.length > 0 && (
        <RadioGroup
          aria-label="Document libraries"
          value={library?.id ?? ""}
          onValueChange={(id) => setLibrary(libs.data.find((l) => l.id === id) ?? null)}
          className="gap-0 divide-y rounded-lg border"
        >
          {libs.data.map((l) => (
            <label key={l.id} htmlFor={`lib-${l.id}`} className="flex cursor-pointer items-center gap-3 px-3 py-3 hover:bg-accent has-[[data-state=checked]]:bg-accent">
              <RadioGroupItem id={`lib-${l.id}`} value={l.id} />
              <Library className="size-4 text-muted-foreground" aria-hidden="true" />
              <span className="min-w-0 flex-1">
                <span className="block text-sm font-medium">{l.displayName}</span>
                <span className="block truncate text-xs text-muted-foreground">{l.rootPath}</span>
              </span>
            </label>
          ))}
        </RadioGroup>
      )}
    </StepFrame>
  );
}

// ---------- 5. Select parent location ----------
export function ParentStep() {
  const site = useWizard((s) => s.site);
  const library = useWizard((s) => s.library);
  const parent = useWizard((s) => s.parent);
  const setParent = useWizard((s) => s.setParent);
  const uploadFlow = useWizard((s) => s.uploadFlow);
  const setUploadFlow = useWizard((s) => s.setUploadFlow);
  const saveConfiguration = useWizard((s) => s.saveConfiguration);
  const goTo = useWizard((s) => s.goTo);

  // Browsing trail below the library root; seeded from an existing selection.
  type TrailEntry = FolderEntry & { rel: string };
  const [trail, setTrail] = useState<TrailEntry[]>(() =>
    parent && parent.relativePath
      ? [{ id: parent.folderId, name: splitRel(parent.relativePath).pop() ?? "", serverPath: parent.serverPath, rel: parent.relativePath }]
      : [],
  );
  const current = trail[trail.length - 1];
  const currentId = current?.id ?? library?.rootFolderId ?? "";
  const relativePath = current?.rel ?? "";
  const serverPath =
    current?.serverPath || (library ? (relativePath ? `${library.rootPath}/${relativePath}` : library.rootPath) : "");

  const folders = useQuery({
    queryKey: ["folders", site?.url, currentId],
    queryFn: async () => {
      try {
        return await SharePointSiteService.listFolders(site!.url, currentId);
      } catch (e) {
        throw toAppError(e, { operation: "ListFolder", action: "read the folders" });
      }
    },
    enabled: Boolean(site && currentId),
    retry: false,
  });

  if (!site || !library) {
    return (
      <StepFrame title="Select parent location" onBack={() => goTo(4)}>
        <p className="text-sm">Select a document library first.</p>
      </StepFrame>
    );
  }

  const selection = { folderId: currentId, relativePath, serverPath };
  const queueProblem = uploadFlow ? validateItemName(uploadFlow.queueFolder.trim()) : null;
  const breadcrumb = [library.displayName, ...splitRel(relativePath)];

  return (
    <StepFrame
      title="Select parent location"
      onBack={() => goTo(4)}
      primary={
        <Button
          disabled={Boolean(queueProblem)}
          onClick={() => {
            if (uploadFlow) setUploadFlow({ queueFolder: uploadFlow.queueFolder.trim() });
            setParent(selection);
            // setParent is synchronous; save reads the latest state.
            saveConfiguration();
          }}
        >
          Save configuration <ArrowRight aria-hidden="true" />
        </Button>
      }
    >
      <nav aria-label="Folder location" className="flex flex-wrap items-center gap-1 text-sm">
        <button type="button" onClick={() => setTrail([])} className="rounded px-1 font-medium hover:underline focus-visible:outline-2 focus-visible:outline-ring" aria-current={trail.length === 0 ? "location" : undefined}>
          {library.displayName}
        </button>
        {trail.map((t, i) => (
          <span key={t.id} className="flex items-center gap-1">
            <ChevronRight className="size-3.5 text-muted-foreground" aria-hidden="true" />
            <button
              type="button"
              onClick={() => setTrail(trail.slice(0, i + 1))}
              className="rounded px-1 hover:underline focus-visible:outline-2 focus-visible:outline-ring"
              aria-current={i === trail.length - 1 ? "location" : undefined}
            >
              {t.name}
            </button>
          </span>
        ))}
      </nav>

      <div>
        <h2 className="mb-2 text-sm font-semibold">Folders in {breadcrumb[breadcrumb.length - 1]}</h2>
        {folders.isPending && <div role="status" aria-label="Loading folders" className="space-y-2">{[0, 1, 2].map((i) => <Skeleton key={i} className="h-10 w-full" />)}</div>}
        {folders.isError && <ErrorPanel title="Folders couldn't be loaded" error={folders.error} action={<Button size="sm" variant="outline" onClick={() => folders.refetch()}>Try again</Button>} />}
        {folders.isSuccess && folders.data.length === 0 && <p className="rounded-lg border px-3 py-3 text-sm">No subfolders here.</p>}
        {folders.isSuccess && folders.data.length > 0 && (
          <ul className="max-h-72 divide-y overflow-auto rounded-lg border">
            {folders.data.map((f) => (
              <li key={f.id}>
                <button
                  type="button"
                  onClick={() => setTrail([...trail, { ...f, rel: relativePath ? `${relativePath}/${f.name}` : f.name }])}
                  className="flex w-full items-center gap-3 px-3 py-2.5 text-left text-sm hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring"
                >
                  <FolderOpen className="size-4 shrink-0 text-muted-foreground" aria-hidden="true" />
                  <span className="min-w-0 flex-1 truncate">{f.name}</span>
                  <span className="sr-only">Open folder</span>
                  <ChevronRight className="size-4 text-muted-foreground" aria-hidden="true" />
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>

      <div className="rounded-lg border bg-card p-4">
        <h2 className="text-sm font-semibold">Selected location</h2>
        <div className="mt-2">
          <PathLine parts={[site.label, ...breadcrumb]} />
        </div>
        <p className="mt-1 text-xs break-all text-muted-foreground">
          {site.url}
          {serverPath}
        </p>
        <p className="mt-2 text-xs text-muted-foreground">{trail.length === 0 ? "Library root" : "Existing folder"}</p>
      </div>

      <div className="space-y-3 rounded-lg border p-4">
        <label htmlFor="use-flow" className="flex items-start gap-3">
          <Checkbox
            id="use-flow"
            checked={Boolean(uploadFlow)}
            onCheckedChange={(v) => setUploadFlow(v === true ? { queueFolder: uploadFlow?.queueFolder ?? DEFAULT_QUEUE_FOLDER } : null)}
            className="mt-0.5"
          />
          <span className="text-sm">
            <span className="block font-medium">Upload PDF, Office and image files through the Power Automate upload flow</span>
            <span className="block text-muted-foreground">Requires the "SPSB Upload" flow set up by your administrator.</span>
          </span>
        </label>
        {uploadFlow && (
          <Field label="Queue folder (in the library root)" htmlFor="queue-folder" error={queueProblem ?? undefined} hint="Must match the folder the flow watches.">
            <Input
              id="queue-folder"
              value={uploadFlow.queueFolder}
              onChange={(e) => setUploadFlow({ queueFolder: e.target.value })}
              aria-invalid={Boolean(queueProblem)}
              aria-describedby={queueProblem ? "queue-folder-error" : "queue-folder-hint"}
              className="max-w-xs font-mono"
            />
          </Field>
        )}
      </div>
    </StepFrame>
  );
}
