// Application state and navigation for the wizard.
// Safe input (selections, project fields, results) is kept in sessionStorage so it
// survives going back and a reload for re-authentication. The selected local
// File is held in memory only and never persisted.
import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import type {
  DocumentLibrary,
  FolderResult,
  ParentLocation,
  ProjectInput,
  SharePointConfiguration,
  SharePointSite,
  UploadFlowSettings,
  UploadResult,
} from "@/models";
import { SharePointConfigurationService } from "@/services/sharepoint-configuration-service";
import { DEFAULT_TEMPLATE_ID } from "@/templates/folder-templates";

export const STEPS = [
  { id: 1, key: "welcome", label: "Welcome", group: "Setup" },
  { id: 2, key: "connection", label: "Microsoft 365 connection", group: "Setup" },
  { id: 3, key: "site", label: "SharePoint site", group: "Setup" },
  { id: 4, key: "library", label: "Document library", group: "Setup" },
  { id: 5, key: "parent", label: "Parent location", group: "Setup" },
  { id: 6, key: "project", label: "Project details", group: "Structure" },
  { id: 7, key: "template", label: "Folder template", group: "Structure" },
  { id: 8, key: "preview", label: "Structure preview", group: "Structure" },
  { id: 9, key: "results", label: "Creation results", group: "Structure" },
  { id: 10, key: "upload", label: "Upload document", group: "Document" },
  { id: 11, key: "upload-confirm", label: "Upload confirmation", group: "Document" },
  { id: 12, key: "upload-result", label: "Upload result", group: "Document" },
] as const;

export type StepId = (typeof STEPS)[number]["id"];

interface WizardState {
  step: StepId;
  config: SharePointConfiguration | null;
  // Setup draft
  site: SharePointSite | null;
  library: DocumentLibrary | null;
  parent: ParentLocation | null;
  uploadFlow: UploadFlowSettings | null;
  // Project
  project: ProjectInput;
  templateId: string;
  folderResults: FolderResult[] | null;
  // Upload
  file: File | null;
  destinationServerPath: string | null;
  targetFileName: string | null;
  uploadResult: UploadResult | null;
  /** In-memory busy flag for running SharePoint mutations (not persisted). */
  busy: boolean;

  setBusy: (busy: boolean) => void;
  goTo: (step: StepId) => void;
  setSite: (site: SharePointSite | null) => void;
  setLibrary: (library: DocumentLibrary | null) => void;
  setParent: (parent: ParentLocation | null) => void;
  setUploadFlow: (uploadFlow: UploadFlowSettings | null) => void;
  saveConfiguration: () => void;
  startReconfigure: () => void;
  setProject: (patch: Partial<ProjectInput>) => void;
  setTemplateId: (id: string) => void;
  setFolderResults: (r: FolderResult[] | null) => void;
  setFile: (f: File | null) => void;
  setDestination: (serverPath: string | null) => void;
  setTargetFileName: (name: string | null) => void;
  setUploadResult: (r: UploadResult | null) => void;
  startNewProject: () => void;
  startNewUpload: () => void;
}

const emptyProject: ProjectInput = { projectNumber: "", projectName: "" };

export const useWizard = create<WizardState>()(
  persist(
    (set, get) => ({
      step: 1,
      config: SharePointConfigurationService.load(),
      site: null,
      library: null,
      parent: null,
      uploadFlow: null,
      project: emptyProject,
      templateId: DEFAULT_TEMPLATE_ID,
      folderResults: null,
      file: null,
      destinationServerPath: null,
      targetFileName: null,
      uploadResult: null,
      busy: false,

      setBusy: (busy) => set({ busy }),
      goTo: (step) => set({ step }),
      setSite: (site) =>
        set((s) => (s.site?.url === site?.url ? { site } : { site, library: null, parent: null })),
      setLibrary: (library) =>
        set((s) => (s.library?.id === library?.id ? { library } : { library, parent: null })),
      setParent: (parent) => set({ parent }),
      setUploadFlow: (uploadFlow) => set({ uploadFlow }),
      saveConfiguration: () => {
        const { site, library, parent, uploadFlow } = get();
        if (!site || !library || !parent) return;
        const config = SharePointConfigurationService.save({ site, library, parent, ...(uploadFlow ? { uploadFlow } : {}) });
        // A new location invalidates results that belonged to the previous one.
        set({ config, step: 6, folderResults: null, destinationServerPath: null, uploadResult: null });
      },
      startReconfigure: () => {
        const { config } = get();
        set({
          step: 2,
          site: config?.site ?? null,
          library: config?.library ?? null,
          parent: config?.parent ?? null,
          uploadFlow: config?.uploadFlow ?? null,
        });
      },
      setProject: (patch) =>
        set((s) => ({ project: { ...s.project, ...patch }, folderResults: null, destinationServerPath: null })),
      setTemplateId: (templateId) => set({ templateId, folderResults: null }),
      setFolderResults: (folderResults) => set({ folderResults }),
      setFile: (file) => set({ file, targetFileName: file?.name ?? null, uploadResult: null }),
      setDestination: (destinationServerPath) => set({ destinationServerPath, uploadResult: null }),
      setTargetFileName: (targetFileName) => set({ targetFileName }),
      setUploadResult: (uploadResult) => set({ uploadResult }),
      startNewProject: () =>
        set({
          step: 6,
          project: emptyProject,
          folderResults: null,
          file: null,
          destinationServerPath: null,
          targetFileName: null,
          uploadResult: null,
        }),
      startNewUpload: () =>
        set({ step: 10, file: null, targetFileName: null, uploadResult: null }),
    }),
    {
      name: "spsb.wizard.v1",
      storage: createJSONStorage(() => sessionStorage),
      // Never persist the File object; persist only safe, non-secret input.
      partialize: (s) => ({
        step: s.step === 11 || s.step === 12 ? 10 : s.step,
        site: s.site,
        library: s.library,
        parent: s.parent,
        uploadFlow: s.uploadFlow,
        project: s.project,
        templateId: s.templateId,
        folderResults: s.folderResults,
        destinationServerPath: s.destinationServerPath,
      }),
      merge: (persisted, current) => ({
        ...current,
        ...(persisted as Partial<WizardState>),
        // Saved configuration always comes from its own service.
        config: SharePointConfigurationService.load(),
      }),
    },
  ),
);
