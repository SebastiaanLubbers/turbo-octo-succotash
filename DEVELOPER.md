# SharePoint Structure Builder — developer guide

A wizard that lets a signed-in Microsoft 365 user pick a SharePoint location,
create a standard project folder structure and upload one document into it.
TypeScript + React, running on the Microsoft Copilot managed app runtime.

## Run and check

```bash
bun install
bun run dev         # local dev server
bun run typecheck   # TypeScript only
bun run lint
```

Open the folder in Visual Studio Code; recommended extensions are listed in
`.vscode/extensions.json`.

## Module layout

| Concern | Location |
| --- | --- |
| Models (`SharePointConfiguration`, `ProjectInput`, `FolderTemplate`, `FolderResult`, `UploadInput`, `UploadResult`) | `src/models/index.ts` |
| **Folder templates** (edit here) | `src/templates/folder-templates.ts` |
| Validation (names, paths, site address, file size) | `src/services/validation.ts` |
| Error handling and result mapping (`AppError`: user message + diagnostic) | `src/services/errors.ts` |
| SharePoint configuration (save/load/clear, recent sites) | `src/services/sharepoint-configuration-service.ts` |
| Site access check, library discovery, folder browsing | `src/services/sharepoint-site-service.ts` |
| Folder plan, existence checks, idempotent creation | `src/services/sharepoint-folder-service.ts` |
| Duplicate detection, verified upload | `src/services/sharepoint-file-service.ts` |
| Hand-off of binary files to the Power Automate upload flow | `src/services/upload-flow-service.ts` (setup: `FLOW-SETUP.md`) |
| Application state and navigation (12 steps) | `src/state/wizard-store.ts`, `src/state/selectors.ts` |
| Wizard host page ("New project") | `src/pages/home.tsx` |
| Existing projects: browse folders/files, add files | `src/pages/projects.tsx`, `src/components/explorer/add-file-panel.tsx` |
| UI components (steps, progress indicator, destination tree, status badges) | `src/components/wizard/` |
| Generated SharePoint connector client (do not edit) | `generated/` |

## Changing or adding a folder template

Edit `FOLDER_TEMPLATES` in `src/templates/folder-templates.ts`. Each template
has a unique `id`, a `name`, a `description` and an ordered `folders` list.
Service logic does not change. Folder names are validated at runtime with the
same rules as user input (`validateItemName`).

## Authentication and data access

* All SharePoint calls go through the app's bound SharePoint Online connector
  (`generated/services/SharePointService.ts`) using the signed-in user's own
  connection and permissions. The app never handles passwords or tokens.
* The connector binding lives in `ms.config.json` (`connectionReferences`) and
  is managed by the App Builder tooling — never edit it by hand.

## Configuration and multi-customer deployment

No tenant ID, hostname, site, library or folder identifier is in the source.
The setup wizard stores the chosen site URL, library and parent folder
(non-secret metadata only) in the browser's `localStorage` under
`spsb.configuration.v1`. In another customer environment, deploy the same
source, let the user create their own SharePoint connection, and run the setup
wizard. "Change configuration" re-runs it at any time.

In-progress wizard input (selections, project fields, results — not the
selected file) is kept in `sessionStorage` so it survives going back and a
reload after re-authentication.

## Behaviour notes

* **Existing projects page:** lists the folders under the configured parent
  folder, browses any project's subfolders and files (name, modified, size,
  link to SharePoint) and adds a file to the open folder with the same upload
  rules as the wizard (confirmation, duplicate rename, verification).

* **Idempotent creation:** every folder is checked (`GetFolderMetadataByPath`)
  before `CreateNewFolder`, and verified after. Existing folders are reported
  as *Already exists*; a failed root marks subfolders *Skipped*; a failed
  subfolder doesn't stop the others. "Retry missing folders" re-runs the same
  plan safely.
* **No overwrite:** the destination name is checked before confirmation and
  again right before upload. On a duplicate the user can upload under the next
  free name (`name (2).ext`) or pick another file/folder. Overwrite is not
  implemented.
* **Binary files via Power Automate:** the app's SharePoint connection sends
  request bodies as UTF-8 text, so PDF/Office/image files can't be written
  directly. When the configuration enables the upload flow, the app writes an
  upload request (JSON with base64 content) as a text file into the queue
  folder; the "SPSB Upload" flow writes the real file through Microsoft Graph
  with credentials from Azure Key Vault and writes a result file back. The app
  then verifies the file and its size. Without the flow, binary files are
  blocked before any SharePoint call and the user is sent to SharePoint.
  See `FLOW-SETUP.md`.
* **Upload verification:** after `CreateFile` the app reads the file's metadata
  and compares the stored size with the local size. On a mismatch it removes
  the file it just created (the name was free moments earlier) and reports
  *Failed*. Uploads are limited to 25 MB (`MAX_UPLOAD_BYTES`).
* **Errors:** `toAppError` maps HTTP 401/403/404/409/429/5xx and policy blocks
  to business messages. Raw connector errors go only to the browser console;
  the UI shows a short diagnostic (operation, HTTP status, time) under
  "Details for support".
* **Site discovery:** the SharePoint connector operations available to the app
  don't include a tenant-wide site search, so the user enters a site address
  (validated by reading its lists) and recently used sites are offered.
