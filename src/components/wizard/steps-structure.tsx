import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { ArrowRight, ExternalLink, FolderPlus, Info, RotateCw, Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Input } from "@/components/ui/input";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Spinner } from "@/components/ui/spinner";
import type { FolderResult } from "@/models";
import { SharePointFolderService } from "@/services/sharepoint-folder-service";
import { AppError, toAppError } from "@/services/errors";
import { buildRootFolderName, validatePlannedPaths, validateProjectInput, hasErrors } from "@/services/validation";
import { FOLDER_TEMPLATES, getTemplate } from "@/templates/folder-templates";
import { sharePointLink, splitRel, usePlan } from "@/state/selectors";
import { useWizard } from "@/state/wizard-store";
import { safeUrl } from "@/lib/safe-url";
import { ErrorPanel } from "./error-panel";
import { PathLine } from "./destination-tree";
import { Field, StepFrame } from "./step-frame";
import { StatusBadge } from "./status";

// ---------- 6. Project details ----------
export function ProjectStep() {
  const project = useWizard((s) => s.project);
  const setProject = useWizard((s) => s.setProject);
  const goTo = useWizard((s) => s.goTo);
  const [touched, setTouched] = useState({ projectNumber: false, projectName: false, submitted: false });
  const errors = validateProjectInput(project);
  const show = (k: "projectNumber" | "projectName") => (touched[k] || touched.submitted ? errors[k] : undefined);
  const rootName = project.projectNumber.trim() && project.projectName.trim() ? buildRootFolderName(project) : "";

  return (
    <StepFrame title="Project details" onBack={() => goTo(1)} backLabel="Welcome">
      <form
        noValidate
        className="space-y-5"
        onSubmit={(e) => {
          e.preventDefault();
          setTouched((t) => ({ ...t, submitted: true }));
          if (!hasErrors(errors)) goTo(7);
        }}
      >
        <div className="grid gap-5 sm:grid-cols-[minmax(0,12rem)_minmax(0,1fr)]">
          <Field label="Project number" htmlFor="project-number" error={show("projectNumber")}>
            <Input
              id="project-number"
              required
              value={project.projectNumber}
              onChange={(e) => setProject({ projectNumber: e.target.value })}
              onBlur={() => setTouched((t) => ({ ...t, projectNumber: true }))}
              aria-invalid={Boolean(show("projectNumber"))}
              aria-describedby={show("projectNumber") ? "project-number-error" : undefined}
              placeholder="P-2026-014"
            />
          </Field>
          <Field label="Project name" htmlFor="project-name" error={show("projectName")}>
            <Input
              id="project-name"
              required
              value={project.projectName}
              onChange={(e) => setProject({ projectName: e.target.value })}
              onBlur={() => setTouched((t) => ({ ...t, projectName: true }))}
              aria-invalid={Boolean(show("projectName"))}
              aria-describedby={show("projectName") ? "project-name-error" : undefined}
              placeholder="Implementatie Finance"
            />
          </Field>
        </div>
        <div className="rounded-lg border bg-card p-4">
          <p className="text-sm font-semibold">Project folder name</p>
          <p className={rootName ? "mt-1 font-mono text-sm" : "mt-1 text-sm text-muted-foreground"}>{rootName || "<project number> - <project name>"}</p>
          {errors.rootFolder && (touched.submitted || rootName) && <p className="mt-1 text-sm text-red-700 dark:text-red-300">{errors.rootFolder}</p>}
        </div>
        <div className="flex justify-end">
          <Button type="submit">
            Continue <ArrowRight aria-hidden="true" />
          </Button>
        </div>
      </form>
    </StepFrame>
  );
}

// ---------- 7. Select folder template ----------
export function TemplateStep() {
  const templateId = useWizard((s) => s.templateId);
  const setTemplateId = useWizard((s) => s.setTemplateId);
  const goTo = useWizard((s) => s.goTo);
  const template = getTemplate(templateId);
  return (
    <StepFrame title="Select folder template" onBack={() => goTo(6)} primary={<Button disabled={!template} onClick={() => goTo(8)}>Continue <ArrowRight aria-hidden="true" /></Button>}>
      <RadioGroup aria-label="Folder templates" value={templateId} onValueChange={setTemplateId} className="gap-3">
        {FOLDER_TEMPLATES.map((t) => (
          <label key={t.id} htmlFor={`tpl-${t.id}`} className="flex cursor-pointer gap-3 rounded-lg border p-4 hover:bg-accent has-[[data-state=checked]]:border-primary">
            <RadioGroupItem id={`tpl-${t.id}`} value={t.id} className="mt-0.5" />
            <span className="min-w-0 flex-1">
              <span className="block font-medium">{t.name}</span>
              <span className="block text-sm text-muted-foreground">{t.description} · {t.folders.length} folders</span>
              <ol className="mt-3 grid gap-x-6 gap-y-1 text-sm sm:grid-cols-2">
                {t.folders.map((f) => (
                  <li key={f} className="font-mono">{f}</li>
                ))}
              </ol>
            </span>
          </label>
        ))}
      </RadioGroup>
    </StepFrame>
  );
}

// ---------- 8. Structure preview ----------
export function PreviewStep({ onCreate }: { onCreate: () => void }) {
  const config = useWizard((s) => s.config);
  const goTo = useWizard((s) => s.goTo);
  const busy = useWizard((s) => s.busy);
  const plan = usePlan();
  const [confirmed, setConfirmed] = useState(false);
  const root = plan?.[0];
  const pathProblem = plan ? validatePlannedPaths(plan.map((p) => p.serverPath)) : null;

  const rootCheck = useQuery({
    queryKey: ["folder-exists", config?.site.url, root?.serverPath],
    queryFn: async () => {
      try {
        return await SharePointFolderService.folderExists(config!.site.url, root!.serverPath);
      } catch (e) {
        throw toAppError(e, { operation: "GetFolderMetadataByPath", action: "check the project folder" });
      }
    },
    enabled: Boolean(config && root),
    retry: false,
    staleTime: 0,
  });

  if (!config || !plan || !root) {
    return (
      <StepFrame title="Structure preview" onBack={() => goTo(6)}>
        <p className="text-sm">Complete the configuration and project details first.</p>
      </StepFrame>
    );
  }

  const blocked = Boolean(pathProblem) || rootCheck.isError || rootCheck.isPending;

  return (
    <StepFrame
      title="Structure preview"
      onBack={() => goTo(7)}
      primary={
        <Button disabled={!confirmed || blocked || busy} onClick={onCreate}>
          <FolderPlus aria-hidden="true" /> Create structure
        </Button>
      }
    >
      <dl className="grid gap-x-6 gap-y-3 rounded-lg border bg-card p-4 text-sm sm:grid-cols-[10rem_minmax(0,1fr)]">
        <dt className="text-muted-foreground">Site</dt>
        <dd className="min-w-0 break-all">{config.site.label} <span className="text-muted-foreground">({config.site.url})</span></dd>
        <dt className="text-muted-foreground">Document library</dt>
        <dd>{config.library.displayName}</dd>
        <dt className="text-muted-foreground">Parent folder</dt>
        <dd>{config.parent.relativePath || "Library root"}</dd>
        <dt className="text-muted-foreground">Project folder</dt>
        <dd className="font-mono font-medium">{root.name}</dd>
        <dt className="text-muted-foreground">Full path</dt>
        <dd className="min-w-0"><PathLine parts={[config.site.label, config.library.displayName, ...splitRel(config.parent.relativePath), root.name]} /></dd>
      </dl>

      <div>
        <h2 className="mb-2 text-sm font-semibold">Subfolders ({plan.length - 1})</h2>
        <ol className="divide-y rounded-lg border text-sm">
          {plan.slice(1).map((f) => (
            <li key={f.name} className="px-3 py-2 font-mono">{f.name}</li>
          ))}
        </ol>
      </div>

      {pathProblem && <ErrorPanel title="Invalid folder name" error={new AppError("validation", pathProblem, "validatePlannedPaths: path length limit")} />}
      {rootCheck.isPending && (
        <p role="status" className="flex items-center gap-2 text-sm"><Spinner /> Checking whether the project folder already exists…</p>
      )}
      {rootCheck.isError && (
        <ErrorPanel title="The destination couldn't be checked" error={rootCheck.error} action={<Button size="sm" variant="outline" onClick={() => rootCheck.refetch()}>Try again</Button>} />
      )}
      {rootCheck.data === "exists" && (
        <div role="status" className="flex gap-3 rounded-lg border border-blue-200 bg-blue-50 p-4 text-sm text-blue-900 dark:border-blue-800 dark:bg-blue-950 dark:text-blue-100">
          <Info className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
          <p>
            <span className="font-medium">Root folder already exists.</span> Existing folders are kept unchanged; only missing subfolders are created. No duplicate folders are made.
          </p>
        </div>
      )}

      <label htmlFor="confirm-create" className="flex items-start gap-3 rounded-lg border p-4">
        <Checkbox id="confirm-create" checked={confirmed} onCheckedChange={(v) => setConfirmed(v === true)} className="mt-0.5" />
        <span className="text-sm">I've reviewed the destination and want to create this folder structure in SharePoint.</span>
      </label>
    </StepFrame>
  );
}

// ---------- 9. Creation progress and results ----------
export function ResultsStep({ onRetry }: { onRetry: () => void }) {
  const config = useWizard((s) => s.config);
  const results = useWizard((s) => s.folderResults);
  const busy = useWizard((s) => s.busy);
  const goTo = useWizard((s) => s.goTo);
  const startNewProject = useWizard((s) => s.startNewProject);
  const plan = usePlan();
  const total = plan?.length ?? 0;
  const done = (results ?? []).filter(Boolean) as FolderResult[];
  const count = (s: FolderResult["status"]) => done.filter((r) => r.status === s).length;
  const rootOk = done[0] && (done[0].status === "Created" || done[0].status === "Already exists");
  const problems = count("Failed") + count("Skipped");
  const finished = !busy && done.length === total && total > 0;
  const partial = finished && problems > 0 && count("Created") + count("Already exists") > 0;
  const uploadTargets = done.filter((r) => !r.isRoot && (r.status === "Created" || r.status === "Already exists")).length;

  return (
    <StepFrame
      title="Creation progress and results"
      onBack={busy ? undefined : () => goTo(8)}
      backLabel="Preview"
      secondary={
        finished && problems > 0 ? (
          <Button variant="outline" onClick={onRetry}>
            <RotateCw aria-hidden="true" /> Retry missing folders
          </Button>
        ) : finished ? (
          <Button variant="outline" onClick={startNewProject}>New project</Button>
        ) : undefined
      }
      primary={
        <Button disabled={!finished || !rootOk || uploadTargets === 0} onClick={() => goTo(10)}>
          <Upload aria-hidden="true" /> Upload document
        </Button>
      }
    >
      <div aria-live="polite" className="space-y-2">
        <div className="flex items-center justify-between text-sm">
          <span className="font-medium">{busy ? "Creating folders…" : finished ? "Finished" : "Not started"}</span>
          <span className="tabular-nums">{done.length} of {total}</span>
        </div>
        <div className="h-2 overflow-hidden rounded-full bg-muted" role="progressbar" aria-label="Folder creation progress" aria-valuemin={0} aria-valuemax={total} aria-valuenow={done.length}>
          <div className="h-full rounded-full bg-primary transition-[width]" style={{ width: total ? `${(done.length / total) * 100}%` : "0%" }} />
        </div>
        {finished && (
          <p className="text-sm">
            {count("Created")} created · {count("Already exists")} already existed · {count("Failed")} failed · {count("Skipped")} skipped
          </p>
        )}
      </div>

      {partial && (
        <div role="alert" className="rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm text-amber-950 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100">
          <p className="font-medium">Partially completed</p>
          <p className="mt-1">Some folders weren't created. Retry creates only what's still missing; folders that exist are left as they are.</p>
        </div>
      )}
      {finished && !rootOk && (
        <div role="alert" className="rounded-lg border border-red-200 bg-red-50 p-4 text-sm text-red-900 dark:border-red-800 dark:bg-red-950 dark:text-red-100">
          <p className="font-medium">Folder creation failed</p>
          <p className="mt-1">The project folder couldn't be created, so no subfolders were created.</p>
        </div>
      )}

      <ol className="divide-y rounded-lg border">
        {(plan ?? []).map((p, i) => {
          const r = results?.[i];
          return (
            <li key={p.serverPath} className="flex flex-wrap items-start gap-x-3 gap-y-1 px-3 py-2.5">
              <span className={p.isRoot ? "min-w-0 flex-1 font-mono text-sm font-semibold" : "min-w-0 flex-1 pl-5 font-mono text-sm"}>{p.name}</span>
              <StatusBadge status={r ? r.status : busy ? "Working" : "Pending"} />
              {r && r.status !== "Created" && (
                <p className="basis-full pl-5 text-xs text-muted-foreground">
                  {r.message}
                  {r.diagnostic && <span className="block font-mono text-[11px]">Support: {r.diagnostic}</span>}
                </p>
              )}
            </li>
          );
        })}
      </ol>

      {finished && rootOk && config && plan && (
        <a
          href={safeUrl(sharePointLink(config.site.url, plan[0].serverPath)) ?? "#"}
          target="_blank"
          rel="noreferrer"
          className="inline-flex items-center gap-1.5 text-sm font-medium text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring"
        >
          Open project folder in SharePoint <ExternalLink className="size-3.5" aria-hidden="true" />
        </a>
      )}
    </StepFrame>
  );
}
