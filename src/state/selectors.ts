import { useMemo } from "react";
import { planStructure, type PlannedFolder } from "@/services/sharepoint-folder-service";
import { getTemplate } from "@/templates/folder-templates";
import { validateProjectInput, hasErrors } from "@/services/validation";
import { useWizard } from "./wizard-store";

/** The planned folders for the current configuration, project and template (null when incomplete). */
export function usePlan(): PlannedFolder[] | null {
  const config = useWizard((s) => s.config);
  const project = useWizard((s) => s.project);
  const templateId = useWizard((s) => s.templateId);
  return useMemo(() => {
    const template = getTemplate(templateId);
    if (!config || !template || hasErrors(validateProjectInput(project))) return null;
    return planStructure(config, project, template);
  }, [config, project, templateId]);
}

/** Absolute link to a site-relative SharePoint path. */
export function sharePointLink(siteUrl: string, serverPath: string): string {
  return `${siteUrl.replace(/\/+$/, "")}${serverPath.split("/").map(encodeURIComponent).join("/")}`;
}

/** Splits a library-relative path into its folder names. */
export function splitRel(rel: string): string[] {
  return rel.split("/").filter(Boolean);
}
