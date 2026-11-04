import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import { after, test } from "node:test";

const require = createRequire(import.meta.url);
const { build } = require(
  require.resolve("esbuild", {
    paths: [dirname(require.resolve("vite/package.json"))],
  }),
);

const bundle = await build({
  stdin: {
    contents:
      'export * from "./src/services/graph-sharepoint.ts"; export { SharePointFileService } from "./src/services/sharepoint-file-service.ts";',
    resolveDir: process.cwd(),
    loader: "ts",
  },
  bundle: true,
  write: false,
  format: "esm",
  platform: "browser",
  plugins: [
    {
      name: "mock-auth",
      setup(plugin) {
        plugin.onResolve({ filter: /browser-auth$/ }, () => ({
          path: "auth",
          namespace: "mock",
        }));
        plugin.onLoad({ filter: /.*/, namespace: "mock" }, () => ({
          contents:
            'export async function graphAccessToken() { return "test-token"; }',
          loader: "js",
        }));
      },
    },
  ],
});
const {
  GraphSharePoint,
  GraphHttpError,
  GraphItemNotFoundError,
  SharePointFileService,
  siteRelativePath,
} = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString("base64")}`
);
const originalFetch = globalThis.fetch;
after(() => {
  globalThis.fetch = originalFetch;
});
const json = (value, status = 200) =>
  new Response(JSON.stringify(value), {
    status,
    headers: { "Content-Type": "application/json" },
  });

test("maps item URLs to site-relative paths and rejects other origins", () => {
  assert.equal(
    siteRelativePath(
      "https://tenant.sharepoint.com/sites/Team",
      "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Project",
    ),
    "/Shared Documents/Project",
  );
  assert.throws(() =>
    siteRelativePath(
      "https://tenant.sharepoint.com/sites/Team",
      "https://other.example/sites/Team/file",
    ),
  );
});

test("follows Graph pages without sending credentials to another origin", async () => {
  const calls = [];
  globalThis.fetch = async (url, init) => {
    calls.push({ url: String(url), headers: init.headers });
    if (calls.length === 1)
      return json({
        value: [{ id: "one" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/next",
      });
    return json({ value: [{ id: "two" }] });
  };
  const result = await GraphSharePoint.children("drive|folder");
  assert.deepEqual(
    result.items.map((item) => item.id),
    ["one", "two"],
  );
  assert.equal(calls[0].headers.Authorization, "Bearer test-token");
  globalThis.fetch = async () =>
    json({ value: [], "@odata.nextLink": "https://evil.example/next" });
  await assert.rejects(
    () => GraphSharePoint.children("drive|folder"),
    /Invalid Microsoft Graph URL/,
  );
});

test("creates folders with conflict behavior set to fail", async () => {
  let request;
  globalThis.fetch = async (url, init) => {
    request = { url: String(url), init };
    return json({ id: "new" }, 201);
  };
  await GraphSharePoint.createFolder("drive", "Parent folder/New folder");
  assert.match(request.url, /root:\/Parent%20folder:\/children$/);
  assert.equal(
    JSON.parse(request.init.body)["@microsoft.graph.conflictBehavior"],
    "fail",
  );
});

test("uploads bytes without forwarding the Graph bearer token to the upload URL", async () => {
  const calls = [];
  globalThis.fetch = async (url, init) => {
    const address = String(url);
    calls.push({ address, init });
    if (address.includes("/sites/tenant.sharepoint.com:"))
      return json({ id: "site" });
    if (address.endsWith("/sites/site/drives"))
      return json({
        value: [
          { id: "drive", name: "Documents", driveType: "documentLibrary" },
        ],
      });
    if (address.endsWith("/drives/drive/root"))
      return json({
        id: "root",
        webUrl: "https://tenant.sharepoint.com/sites/Team/Shared%20Documents",
      });
    if (address.includes("createUploadSession"))
      return json({ uploadUrl: "https://upload.example/session?token=secret" });
    return json({ id: "uploaded", size: 3 }, 201);
  };
  const file = new File([new Uint8Array([0, 128, 255])], "file.bin");
  const item = await GraphSharePoint.upload(
    "https://tenant.sharepoint.com/sites/Team",
    "/Shared Documents/Folder/file.bin",
    file,
  );
  const session = calls.find((call) =>
    call.address.includes("createUploadSession"),
  );
  const upload = calls.at(-1);
  assert.equal(
    JSON.parse(session.init.body).item["@microsoft.graph.conflictBehavior"],
    "fail",
  );
  assert.equal(upload.init.headers.Authorization, undefined);
  assert.equal(upload.init.headers["Content-Range"], "bytes 0-2/3");
  assert.deepEqual(
    new Uint8Array(await upload.init.body.arrayBuffer()),
    new Uint8Array([0, 128, 255]),
  );
  assert.equal(item.id, "uploaded");
});

test("preserves Graph conflict status", async () => {
  globalThis.fetch = async () => json({}, 409);
  await assert.rejects(
    () => GraphSharePoint.createFolder("drive", "Existing"),
    (error) => error instanceof GraphHttpError && error.status === 409,
  );
});

test("distinguishes a missing item from a missing site", async () => {
  globalThis.fetch = async (url) => {
    const address = String(url);
    if (address.includes("/sites/tenant.sharepoint.com:"))
      return json({ id: "site" });
    if (address.endsWith("/sites/site/drives"))
      return json({
        value: [
          { id: "drive", name: "Documents", driveType: "documentLibrary" },
        ],
      });
    if (address.endsWith("/drives/drive/root"))
      return json({
        id: "root",
        webUrl: "https://tenant.sharepoint.com/sites/Team/Documents",
      });
    return json({}, 404);
  };
  await assert.rejects(
    () =>
      GraphSharePoint.itemAtPath(
        "https://tenant.sharepoint.com/sites/Team",
        "/Documents/Missing",
      ),
    GraphItemNotFoundError,
  );
  globalThis.fetch = async () => json({}, 404);
  await assert.rejects(
    () =>
      GraphSharePoint.itemAtPath(
        "https://tenant.sharepoint.com/sites/Team",
        "/Documents/Missing",
      ),
    (error) =>
      error instanceof GraphHttpError &&
      !(error instanceof GraphItemNotFoundError),
  );
});

test("uploads a text file after a missing-file check and verifies it", async () => {
  const stages = [];
  let uploaded = false;
  globalThis.fetch = async (url) => {
    const address = String(url);
    if (address.includes("/sites/tenant.sharepoint.com:"))
      return json({ id: "site" });
    if (address.endsWith("/sites/site/drives"))
      return json({
        value: [
          { id: "drive", name: "Documents", driveType: "documentLibrary" },
        ],
      });
    if (address.endsWith("/drives/drive/root"))
      return json({
        id: "root",
        webUrl: "https://tenant.sharepoint.com/sites/Team/Documents",
      });
    if (address.includes("createUploadSession"))
      return json({ uploadUrl: "https://upload.example/session" });
    if (address.startsWith("https://upload.example/")) {
      uploaded = true;
      return json({ id: "uploaded", size: 3 }, 201);
    }
    return uploaded ? json({ id: "uploaded", size: 3 }) : json({}, 404);
  };
  const file = new File(["abc"], "notes.txt", { type: "text/plain" });
  const result = await SharePointFileService.upload(
    "https://tenant.sharepoint.com/sites/Team",
    {
      file,
      folderServerPath: "/Documents/Folder",
      folderRelativePath: "Folder",
      targetFileName: file.name,
    },
    (stage) => stages.push(stage),
  );
  assert.equal(result.status, "Uploaded");
  assert.deepEqual(stages, ["checking", "uploading", "verifying", "done"]);
});

test("does not upload when the destination lookup fails", async () => {
  const originalError = console.error;
  console.error = () => {};
  let uploadRequested = false;
  try {
    globalThis.fetch = async (url) => {
      if (String(url).includes("createUploadSession")) uploadRequested = true;
      return json({}, 403);
    };
    const file = new File(["abc"], "notes.txt");
    const result = await SharePointFileService.upload(
      "https://tenant.sharepoint.com/sites/Team",
      {
        file,
        folderServerPath: "/Documents/Folder",
        folderRelativePath: "Folder",
        targetFileName: file.name,
      },
      () => {},
    );
    assert.equal(result.status, "Failed");
    assert.match(result.message, /permission/);
    assert.equal(uploadRequested, false);
  } finally {
    console.error = originalError;
  }
});
