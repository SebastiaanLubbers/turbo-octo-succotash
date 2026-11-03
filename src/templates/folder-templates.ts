import type { FolderTemplate } from "@/models";

// Folder templates — maintain the structures here. SharePoint service logic does
// not depend on the contents of this file: add a template object to the array
// (unique `id`) or change a `folders` list, and the wizard picks it up.
// Folder names are validated at runtime with the same rules as user input.
export const FOLDER_TEMPLATES: FolderTemplate[] = [
  {
    id: "project",
    name: "Project",
    description: "Standard project structure",
    folders: [
      "01. Projectmanagement",
      "02. Contracten",
      "03. Analyse",
      "04. Ontwerp",
      "05. Realisatie",
      "06. Testen",
      "07. Oplevering",
      "08. Beheer",
    ],
  },
];

export const DEFAULT_TEMPLATE_ID = "project";

export function getTemplate(id: string): FolderTemplate | undefined {
  return FOLDER_TEMPLATES.find((t) => t.id === id);
}
