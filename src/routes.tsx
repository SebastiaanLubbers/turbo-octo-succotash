import type { ReactElement } from "react";
import { FolderPlus, FolderTree, type LucideIcon } from "lucide-react";
import { HomePage } from "@/pages/home";
import { ProjectsPage } from "@/pages/projects";

export interface AppRoute {
  /** "/" is the index route; others are paths under the shell. */
  path: string;
  element: ReactElement;
  /** Optional nav metadata — routes with a label appear in the top bar. */
  label?: string;
  icon?: LucideIcon;
}

// Single source of truth for routes. Add a page = add ONE entry here.
export const routes: AppRoute[] = [
  { path: "/", element: <HomePage />, label: "New project", icon: FolderPlus },
  { path: "/projects", element: <ProjectsPage />, label: "Existing projects", icon: FolderTree },
];
