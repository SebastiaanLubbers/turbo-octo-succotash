# "SPSB Upload" Power Automate flow — setup guide

> Historical reference only: the app now uploads all supported files directly
> through Microsoft Graph and no longer creates or reads upload queue files.
> Existing flows and queue contents are not removed automatically. The setup
> steps below describe the retired integration and should not be used to
> configure the current app.

This flow lets SharePoint Structure Builder upload **any** file type (PDF,
Office, images). The app writes an upload request as a JSON text file into a
queue folder; the flow decodes it and writes the real file through Microsoft
Graph with an app registration whose **client secret (or certificate) is kept in
Azure Key Vault**. The app never holds a flow URL, secret or token.

```
App ──CreateFile (text)──▶ <library>/_UploadQueue/<id>.request.json
                                   │  (trigger)
                                   ▼
                    Flow: Key Vault secret → Graph PUT file content
                                   │
App ◀──reads/deletes── <library>/_UploadQueue/<id>.result.json
```

## Request and result contract

Request file `<requestId>.request.json` (written by the app):

```json
{
  "format": "spsb.upload-request",
  "version": 1,
  "requestId": "6c1f…",
  "requestedAt": "2026-11-03T18:00:00Z",
  "siteUrl": "https://<tenant>.sharepoint.com/sites/<site>",
  "libraryId": "<library (list) id>",
  "targetFolder": "Projects/P-001 - Example/02. Contracten",
  "fileName": "Contract.pdf",
  "contentType": "application/pdf",
  "size": 1117437,
  "contentBase64": "JVBERi0xLjcK…"
}
```

Result file `<requestId>.result.json` (written by the flow, in the same folder):

```json
{ "status": "Uploaded", "message": "OK" }
{ "status": "Failed", "message": "A file with this name already exists." }
```

The app waits up to 3 minutes for the result, then checks the file exists at the
destination with the same size before reporting success. Without a result it
reports **Pending** (never success).

## 1. Azure / Entra prerequisites (administrator)

1. **App registration** in Microsoft Entra ID, e.g. `SPSB Upload Flow`.
   - API permission: Microsoft Graph → **Application** → `Sites.Selected`
     (grant admin consent). This gives access to _no_ site until granted below.
   - Create a **client secret** (or upload a certificate). Copy the value once.
2. **Grant the app write access to the project site only** (a SharePoint/Graph
   admin runs this once per site, e.g. in Graph Explorer as admin):

   ```http
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   { "roles": ["write"],
     "grantedToIdentities": [ { "application": { "id": "<app client id>", "displayName": "SPSB Upload Flow" } } ] }
   ```

   (`{site-id}` from `GET https://graph.microsoft.com/v1.0/sites/<tenant>.sharepoint.com:/sites/<site>`.)

3. **Azure Key Vault**: store the secret as `spsb-upload-client-secret`.
   Give the account used by the flow's Key Vault connection the
   **Key Vault Secrets User** role on the vault (RBAC) or a _Get secret_ access policy.
4. **Licensing**: the _HTTP_ and _Azure Key Vault_ connectors are premium; the
   flow owner needs a Power Automate Premium (or per-flow) licence.
5. **DLP**: allow SharePoint, HTTP and Azure Key Vault together in the
   environment's data policy.

## 2. Queue folder

In the project library's root create a folder `_UploadQueue` (or the name
entered in the app). The app also creates it on first use if missing.
Users need Contribute access to it; consider hiding it from views and adding a
retention/cleanup rule. Request files contain file content until the flow
deletes them.

## 3. Build the flow (Power Automate → Create → Automated cloud flow)

Name: **SPSB Upload**. Use a service account (not a personal account) as owner
and for the SharePoint connection.

| #     | Action                                                                                                | Settings                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ----- | ----------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1     | **SharePoint – When a file is created in a folder**                                                   | Site Address: the project site. Folder Id: `/<library path>/_UploadQueue`. _Settings → Trigger conditions_: `@endsWith(triggerOutputs()?['headers']?['x-ms-file-name'], '.request.json')`. _Concurrency control_: on, degree 5.                                                                                                                                                                                           |
| 2     | **Compose** – rename to `Request`                                                                     | Inputs (expression): `json(base64ToString(triggerBody()?['$content']))`                                                                                                                                                                                                                                                                                                                                                   |
| 3     | **Initialize variable** `AllowedRoot` (String)                                                        | The library-relative parent folder configured in the app, e.g. `Projects` (empty only if the app uses the library root).                                                                                                                                                                                                                                                                                                  |
| 4     | **Initialize variable** `ResultJson` (String)                                                         | `{"status":"Failed","message":"The upload flow couldn't store the file."}`                                                                                                                                                                                                                                                                                                                                                |
| 5     | **Condition** – `Allowed destination`                                                                 | Expression is **true**: `and(equals(outputs('Request')?['format'],'spsb.upload-request'), startsWith(toLower(outputs('Request')?['targetFolder']), toLower(variables('AllowedRoot'))), not(contains(outputs('Request')?['targetFolder'],'..')), not(contains(outputs('Request')?['fileName'],'/')))`                                                                                                                      |
| 5-No  | **Set variable** `ResultJson`                                                                         | `{"status":"Failed","message":"This destination isn't allowed by the upload flow."}`                                                                                                                                                                                                                                                                                                                                      |
| 5-Yes | **Scope** – rename to `Upload`                                                                        | contains 6–9                                                                                                                                                                                                                                                                                                                                                                                                              |
| 6     | **Azure Key Vault – Get secret**                                                                      | Name: `spsb-upload-client-secret`. _Settings → Secure outputs_: **On**.                                                                                                                                                                                                                                                                                                                                                   |
| 7     | **HTTP** – rename to `Get site`                                                                       | GET `https://graph.microsoft.com/v1.0/sites/@{uriHost(outputs('Request')?['siteUrl'])}:@{uriPath(outputs('Request')?['siteUrl'])}`. Authentication: **Active Directory OAuth** — Authority `https://login.microsoftonline.com`, Tenant `<tenant id>`, Audience `https://graph.microsoft.com`, Client ID `<app client id>`, Credential type **Secret**, Secret `@{body('Get_secret')?['value']}`. _Secure inputs_: **On**. |
| 8     | **HTTP** – rename to `Get drive`                                                                      | GET `https://graph.microsoft.com/v1.0/sites/@{body('Get_site')?['id']}/lists/@{outputs('Request')?['libraryId']}/drive` — same authentication, _Secure inputs_ On.                                                                                                                                                                                                                                                        |
| 9     | **HTTP** – rename to `Put file`                                                                       | PUT `https://graph.microsoft.com/v1.0/drives/@{body('Get_drive')?['id']}/root:/@{replace(uriComponent(concat(outputs('Request')?['targetFolder'],'/',outputs('Request')?['fileName'])),'%2F','/')}:/content?@microsoft.graph.conflictBehavior=fail`. Header `Content-Type: application/octet-stream`. Body (expression): `base64ToBinary(outputs('Request')?['contentBase64'])`. Same authentication, _Secure inputs_ On. |
| 10    | **Set variable** `ResultJson` (after the Scope, _Configure run after_: **is successful**)             | `{"status":"Uploaded","message":"OK"}`                                                                                                                                                                                                                                                                                                                                                                                    |
| 11    | **Set variable** `ResultJson` (after the Scope, _Configure run after_: **has failed, has timed out**) | `{"status":"Failed","message":"@{if(equals(outputs('Put_file')?['statusCode'],409),'A file with this name already exists.','The upload flow couldn''t store the file. Ask your administrator to check the flow run.')}"}`                                                                                                                                                                                                 |
| 12    | **SharePoint – Create file** (after the Condition, run after _is successful_)                         | Folder Path: `/<library path>/_UploadQueue`. File Name: `@{outputs('Request')?['requestId']}.result.json`. File Content: `@{variables('ResultJson')}`                                                                                                                                                                                                                                                                     |
| 13    | **SharePoint – Delete file** (run after 12: _is successful, has failed_)                              | File Identifier: `@{triggerOutputs()?['headers']?['x-ms-file-id']}` (removes the request with the file content).                                                                                                                                                                                                                                                                                                          |

Notes

- If action 2 fails because the trigger body is already JSON, use `triggerBody()` instead.
- Graph's simple upload handles up to 250 MB; the app limits uploads to 25 MB.
- Keep _Secure inputs/outputs_ on for steps 6–9 so the secret and file content
  don't appear in run history.
- For a certificate instead of a secret: store it in Key Vault, choose credential
  type **Certificate** in the HTTP actions, and pass the PFX and password from Key Vault.

## 4. Enable it in the app

In the setup wizard, step **Select parent location**, tick **Upload PDF, Office
and image files through the Power Automate upload flow** and enter the queue
folder name (default `_UploadQueue`). Save the configuration. Text files still
upload directly; other files go through the flow.

## 5. Test

1. Upload a small PDF to a project subfolder from the app.
2. Expect "Upload complete"; the file opens in SharePoint.
3. Upload the same file again → the app offers a new name; nothing is overwritten.
4. In the flow's run history, confirm that inputs/outputs of steps 6–9 are hidden.
