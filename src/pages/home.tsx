import { useEffect } from "react";
import type { FolderResult } from "@/models";
import { queryClient } from "@/lib/query-client";
import { DestinationTree, type TreeInput } from "@/components/wizard/destination-tree";
import { StepIndicator } from "@/components/wizard/step-indicator";
import { ConnectionStep, LibraryStep, ParentStep, SiteStep, WelcomeStep } from "@/components/wizard/steps-setup";
import { PreviewStep, ProjectStep, ResultsStep, TemplateStep } from "@/components/wizard/steps-structure";
import { UploadConfirmStep, UploadResultStep, UploadStep } from "@/components/wizard/steps-upload";
import { SharePointFolderService } from "@/services/sharepoint-folder-service";
import { buildRootFolderName } from "@/services/validation";
import { getTemplate } from "@/templates/folder-templates";
import { splitRel, usePlan } from "@/state/selectors";
import { type StepId, useWizard } from "@/state/wizard-store";

export function HomePage() {
  const step = useWizard((s) => s.step);
  const busy = useWizard((s) => s.busy);
  const config = useWizard((s) => s.config);
  const site = useWizard((s) => s.site);
  const library = useWizard((s) => s.library);
  const parent = useWizard((s) => s.parent);
  const project = useWizard((s) => s.project);
  const templateId = useWizard((s) => s.templateId);
  const folderResults = useWizard((s) => s.folderResults);
  const destination = useWizard((s) => s.destinationServerPath);
  const file = useWizard((s) => s.file);
  const targetFileName = useWizard((s) => s.targetFileName);
  const goTo = useWizard((s) => s.goTo);
  const startReconfigure = useWizard((s) => s.startReconfigure);
  const plan = usePlan();

  // Move focus to the new step's heading so keyboard and screen-reader users land on it.
  useEffect(() => {
    document.getElementById("step-title")?.focus({ preventScroll: false });
  }, [step]);

  // Without saved configuration the app always starts in the setup wizard.
  useEffect(() => {
    if (!config && step >= 6) goTo(1);
  }, [config, step, goTo]);

  async function runCreation() {
    const { config: cfg, setBusy, setFolderResults } = useWizard.getState();
    if (!cfg || !plan || useWizard.getState().busy) return;
    setBusy(true);
    setFolderResults([]);
    goTo(9);
    const live: FolderResult[] = [];
    try {
      await SharePointFolderService.createStructure(cfg, plan, (i, r) => {
        live[i] = r;
        setFolderResults([...live]);
      });
    } finally {
      setBusy(false);
      // Existing-projects views re-read SharePoint after a change.
      void queryClient.invalidateQueries({ queryKey: ["contents"] });
    }
  }

  const rootOk = folderResults?.[0] && (folderResults[0].status === "Created" || folderResults[0].status === "Already exists");
  function canVisit(id: StepId): boolean {
    if (busy || id >= step) return false;
    if (id >= 2 && id <= 5) return true;
    if (id >= 6 && id <= 8) return Boolean(config);
    if (id === 9) return Boolean(folderResults?.length);
    if (id === 10) return Boolean(rootOk);
    return id === 1;
  }
  function visit(id: StepId) {
    if (id >= 2 && id <= 5 && step >= 6 && !site) {
      startReconfigure();
    }
    goTo(id);
  }

  // Live destination tree beside every step after the welcome.
  const inSetup = step <= 5;
  const tplFolders = getTemplate(templateId)?.folders ?? [];
  const rootName = project.projectNumber.trim() && project.projectName.trim() ? buildRootFolderName(project) : "";
  const destIndex = plan ? plan.slice(1).findIndex((p) => p.serverPath === destination) : -1;
  const tree: TreeInput = inSetup
    ? {
        siteLabel: site?.label,
        siteUrl: site?.url,
        libraryName: library?.displayName,
        parentSegments: parent ? splitRel(parent.relativePath) : [],
      }
    : {
        siteLabel: config?.site.label,
        siteUrl: config?.site.url,
        libraryName: config?.library.displayName,
        parentSegments: config ? splitRel(config.parent.relativePath) : [],
        rootName,
        subfolders: tplFolders,
        results: step >= 9 ? folderResults : null,
        running: busy && step === 9,
        fileFolderIndex: step >= 10 && destIndex >= 0 ? destIndex : undefined,
        fileName: step >= 10 ? (targetFileName ?? file?.name ?? undefined) : undefined,
      };

  const body = (() => {
    switch (step) {
      case 1: return <WelcomeStep />;
      case 2: return <ConnectionStep />;
      case 3: return <SiteStep />;
      case 4: return <LibraryStep />;
      case 5: return <ParentStep />;
      case 6: return <ProjectStep />;
      case 7: return <TemplateStep />;
      case 8: return <PreviewStep onCreate={() => void runCreation()} />;
      case 9: return <ResultsStep onRetry={() => void runCreation()} />;
      case 10: return <UploadStep />;
      case 11: return <UploadConfirmStep />;
      case 12: return <UploadResultStep />;
    }
  })();

  return (
    <div className="mx-auto grid w-full max-w-[1400px] gap-8 px-4 py-6 md:grid-cols-[15rem_minmax(0,1fr)] md:px-6 lg:py-8 xl:grid-cols-[15rem_minmax(0,1fr)_22rem]">
      <aside className="md:sticky md:top-20 md:self-start">
        <StepIndicator canVisit={(id) => canVisit(id)} onVisit={visit} />
      </aside>
      <div id="wizard-content" tabIndex={-1} className="min-w-0 outline-none">
        {body}
      </div>
      {step > 1 && (
        <aside aria-label="Destination overview" className="min-w-0 md:col-start-2 xl:sticky xl:top-20 xl:col-start-3 xl:self-start">
          <DestinationTree input={tree} caption={inSetup ? "Selected location" : "Destination"} />
        </aside>
      )}
    </div>
  );
}
