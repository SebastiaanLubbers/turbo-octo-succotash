import { FileText, Folder, FolderOpen, Globe, Library } from "lucide-react";
import type { FolderResult } from "@/models";
import { cn } from "@/lib/utils";
import { StatusBadge } from "./status";

export interface TreeInput {
  siteLabel?: string;
  siteUrl?: string;
  libraryName?: string;
  parentSegments?: string[];
  rootName?: string;
  subfolders?: string[];
  results?: (FolderResult | undefined)[] | null;
  /** Index into subfolders that receives the file. */
  fileFolderIndex?: number;
  fileName?: string;
  /** Show the "pending" state instead of nothing when results are running. */
  running?: boolean;
}

function Row({
  depth,
  icon: Icon,
  label,
  sub,
  emphasis,
  placeholder,
  trailing,
}: {
  depth: number;
  icon: typeof Folder;
  label: string;
  sub?: string;
  emphasis?: boolean;
  placeholder?: boolean;
  trailing?: React.ReactNode;
}) {
  return (
    <li className="relative" style={{ paddingLeft: depth * 18 }}>
      {depth > 0 && (
        <span aria-hidden="true" className="absolute top-0 h-1/2 w-3 border-b border-l border-border" style={{ left: depth * 18 - 11 }} />
      )}
      <div className={cn("flex min-w-0 items-center gap-2 rounded-md px-1.5 py-1", emphasis && "bg-accent")}>
        <Icon className={cn("size-4 shrink-0", emphasis ? "text-primary" : "text-muted-foreground")} aria-hidden="true" />
        <span className={cn("min-w-0 truncate text-sm", emphasis && "font-semibold", placeholder && "italic text-muted-foreground")} title={label}>
          {label}
        </span>
        {sub && <span className="hidden min-w-0 truncate text-xs text-muted-foreground xl:inline">{sub}</span>}
        {trailing && <span className="ml-auto pl-2">{trailing}</span>}
      </div>
    </li>
  );
}

/** The destination as a folder tree — the working visual for setup, preview, results and upload. */
export function DestinationTree({ input, caption = "Destination" }: { input: TreeInput; caption?: string }) {
  const parents = input.parentSegments ?? [];
  let depth = 0;
  const rows: React.ReactNode[] = [];
  rows.push(
    <Row key="site" depth={depth++} icon={Globe} label={input.siteLabel ?? "Site not selected"} sub={input.siteUrl} placeholder={!input.siteLabel} />,
  );
  rows.push(<Row key="lib" depth={depth++} icon={Library} label={input.libraryName ?? "Library not selected"} placeholder={!input.libraryName} />);
  parents.forEach((p, i) => rows.push(<Row key={`p${i}`} depth={depth++} icon={FolderOpen} label={p} />));
  if (input.rootName !== undefined) {
    const rootResult = input.results?.[0];
    rows.push(
      <Row
        key="root"
        depth={depth++}
        icon={FolderOpen}
        label={input.rootName || "<project number> - <project name>"}
        placeholder={!input.rootName}
        emphasis
        trailing={rootResult ? <StatusBadge status={rootResult.status} /> : input.running ? <StatusBadge status="Pending" /> : undefined}
      />,
    );
    (input.subfolders ?? []).forEach((f, i) => {
      const r = input.results?.[i + 1];
      rows.push(
        <Row
          key={`s${i}`}
          depth={depth}
          icon={Folder}
          label={f}
          emphasis={input.fileFolderIndex === i}
          trailing={r ? <StatusBadge status={r.status} /> : input.running ? <StatusBadge status="Pending" /> : undefined}
        />,
      );
      if (input.fileFolderIndex === i && input.fileName) {
        rows.push(<Row key="file" depth={depth + 1} icon={FileText} label={input.fileName} emphasis />);
      }
    });
  }
  return (
    <figure aria-labelledby="destination-caption" className="rounded-lg border bg-card p-4">
      <figcaption id="destination-caption" className="mb-3 text-sm font-semibold">
        {caption}
      </figcaption>
      <ul className="space-y-0.5" aria-label={`${caption} folder tree`}>{rows}</ul>
    </figure>
  );
}

/** Full path as one readable line, e.g. Site › Library › Folder › File. */
export function PathLine({ parts }: { parts: string[] }) {
  return (
    <p className="text-sm break-words">
      {parts.map((p, i) => (
        <span key={i}>
          {i > 0 && <span className="mx-1.5 text-muted-foreground" aria-hidden="true">›</span>}
          {i > 0 && <span className="sr-only">, then </span>}
          <span className={cn(i === parts.length - 1 && "font-medium")}>{p}</span>
        </span>
      ))}
    </p>
  );
}
