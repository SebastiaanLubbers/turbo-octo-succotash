# SharePoint Structure Builder — developer guide

A wizard that lets a signed-in Microsoft 365 user pick a SharePoint location,
create a standard project folder structure and upload one document into it.
TypeScript + React, using browser sign-in and Microsoft Graph.

## Run and check

```bash
bun install
bun run dev         # local dev server
bun run typecheck   # TypeScript only
bun run lint
node --test tests/graph-sharepoint.test.mjs
```

Open the folder in Visual Studio Code; recommended extensions are listed in
`.vscode/extensions.json`.

## Module layout

| Concern                                                                                                             | Location                                                               |
| ------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| Models (`SharePointConfiguration`, `ProjectInput`, `FolderTemplate`, `FolderResult`, `UploadInput`, `UploadResult`) | `src/models/index.ts`                                                  |
| **Folder templates** (edit here)                                                                                    | `src/templates/folder-templates.ts`                                    |
| Validation (names, paths, site address, file size)                                                                  | `src/services/validation.ts`                                           |
| Error handling and result mapping (`AppError`: user message + diagnostic)                                           | `src/services/errors.ts`                                               |
| SharePoint configuration (save/load/clear, recent sites)                                                            | `src/services/sharepoint-configuration-service.ts`                     |
| Site access check, library discovery, folder browsing                                                               | `src/services/sharepoint-site-service.ts`                              |
| Folder plan, existence checks, idempotent creation                                                                  | `src/services/sharepoint-folder-service.ts`                            |
| Duplicate detection, verified upload                                                                                | `src/services/sharepoint-file-service.ts`                              |
| Application state and navigation (12 steps)                                                                         | `src/state/wizard-store.ts`, `src/state/selectors.ts`                  |
| Wizard host page ("New project")                                                                                    | `src/pages/home.tsx`                                                   |
| Existing projects: browse folders/files, add files                                                                  | `src/pages/projects.tsx`, `src/components/explorer/add-file-panel.tsx` |
| UI components (steps, progress indicator, destination tree, status badges)                                          | `src/components/wizard/`                                               |
| Generated SharePoint connector client (do not edit)                                                                 | `generated/`                                                           |

## Changing or adding a folder template

Edit `FOLDER_TEMPLATES` in `src/templates/folder-templates.ts`. Each template
has a unique `id`, a `name`, a `description` and an ordered `folders` list.
Service logic does not change. Folder names are validated at runtime with the
same rules as user input (`validateItemName`).

## Authentication and data access

- Register a Microsoft Entra application for the intended tenant or tenants.
  Add a **Single-page application** redirect URI for the exact deployed callback,
  for example `http://localhost:5173/auth-callback.html` locally or
  `https://OWNER.github.io/REPOSITORY/auth-callback.html` on GitHub Pages.
- Add Microsoft Graph delegated permissions `User.Read` and
  `Sites.ReadWrite.All`; obtain tenant admin consent when required. The latter
  is broad: it grants the app delegated read/write access to the site collections
  the user can access. SharePoint permissions and organizational policies still apply.
- Set `VITE_ENTRA_CLIENT_ID` to the application (client) ID and
  `VITE_ENTRA_TENANT_ID` to the directory (tenant) ID before building or running.
  If the tenant ID is omitted, `organizations` is used for a multitenant app.
  These settings are public build-time values, not secrets. Never put a client
  secret in the frontend. Rebuild when changing them.
- MSAL uses a popup with a passive same-origin callback and sessionStorage;
  popup blocking, cross-origin opener policies, and embedded browser restrictions
  can prevent sign-in. Sign out clears this app's local MSAL cache; it does not
  sign the user out of Microsoft globally.
- Microsoft Graph performs the SharePoint requests with delegated access tokens.
  The managed connector binding in `ms.config.json` is not used by these workflows.
  The former Power Automate upload flow is no longer used; existing flows and
  queue files are not changed by the app.

## GitHub Pages

The Pages workflow builds and deploys a repository project site when a commit
reaches `main`. It can also be started manually from the Actions tab.

1. In repository **Settings > Pages**, set the source to **GitHub Actions**.
2. Under **Settings > Secrets and variables > Actions > Secrets**, add
   `VITE_ENTRA_CLIENT_ID` and optionally `VITE_ENTRA_TENANT_ID`. These IDs are
   included in the public build even when stored as GitHub secrets; never use
   an Entra client secret here. The tenant defaults to `organizations` when
   omitted.
3. Register `https://OWNER.github.io/REPOSITORY/auth-callback.html` as a
   **Single-page application** redirect URI in Entra, and grant the delegated
   Microsoft Graph permissions described above.
4. Push to `main` or run the workflow manually. The workflow builds with
   `VITE_BASE=/REPOSITORY/` and publishes `dist` to the project site.

The app uses browser history routes. Navigation from the root works, but direct
requests or refreshes at `/projects` return a Pages 404 without a separate SPA
fallback. Confirm your organization permits publishing the frontend publicly.

## Configuration and multi-customer deployment

The app registration is supplied at build time; no site, library or folder is hard-coded.
The setup wizard stores the chosen site URL, library and parent folder
(non-secret metadata only) in the browser's `localStorage` under
`spsb.graph.configuration.v1`. Graph uses different library and folder IDs from
the managed connector, so existing users must select their location again.
In another customer environment, deploy with an app registration approved by
that tenant and run the setup wizard. "Change configuration" re-runs it at any time.

In-progress wizard input (selections, project fields, results — not the
selected file) is kept in `sessionStorage` so it survives going back and a
reload after re-authentication.

## Behaviour notes

- **Existing projects page:** lists the folders under the configured parent
  folder, browses any project's subfolders and files (name, modified, size,
  link to SharePoint) and adds a file to the open folder with the same upload
  rules as the wizard (confirmation, duplicate rename, verification).

- **Idempotent creation:** every folder is checked (`GetFolderMetadataByPath`)
  before `CreateNewFolder`, and verified after. Existing folders are reported
  as _Already exists_; a failed root marks subfolders _Skipped_; a failed
  subfolder doesn't stop the others. "Retry missing folders" re-runs the same
  plan safely.
- **No overwrite:** the destination name is checked before confirmation and
  again right before upload. On a duplicate the user can upload under the next
  free name (`name (2).ext`) or pick another file/folder. Overwrite is not
  implemented.
- **Upload verification:** binary and text files are uploaded using Graph upload
  sessions with conflict behavior `fail`. The app checks the returned item ID
  and stored size against the local file; it never deletes a mismatched item.
  Uploads are limited to 25 MB (`MAX_UPLOAD_BYTES`). A failed or interrupted
  upload may leave an uncommitted upload session until Microsoft expires it.
- **Errors:** `toAppError` maps HTTP 401/403/404/409/429/5xx and policy blocks
  to business messages. Raw connector errors go only to the browser console;
  the UI shows a short diagnostic (operation, HTTP status, time) under
  "Details for support".
- **Site discovery:** the user enters a site address (validated through Graph)
  and recently used sites are offered.
