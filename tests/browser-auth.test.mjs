import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import { test } from "node:test";

const require = createRequire(import.meta.url);
const { build } = require(require.resolve("esbuild", { paths: [dirname(require.resolve("vite/package.json"))] }));
const bundle = await build({
  stdin: {
    contents: 'export { signInErrorMessage } from "./src/services/browser-auth.ts"; export { AuthError } from "@azure/msal-browser";',
    resolveDir: process.cwd(),
    loader: "ts",
  },
  bundle: true,
  write: false,
  format: "esm",
  platform: "browser",
  define: { "import.meta.env": JSON.stringify({ BASE_URL: "/" }) },
});
globalThis.window = { location: { origin: "http://localhost:5173" } };
const { signInErrorMessage, AuthError } = await import(`data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString("base64")}`);

test("shows the configured callback for a redirect mismatch without exposing the raw response", () => {
  const error = new AuthError("invalid_request", "", "AADSTS50011: private-response-value");
  const message = signInErrorMessage(error);
  assert.match(message, /AADSTS50011/);
  assert.match(message, /http:\/\/localhost:5173\/auth-callback\.html/);
  assert.doesNotMatch(message, /private-response-value/);
});

test("gives actionable popup guidance and limits untrusted error codes", () => {
  assert.match(signInErrorMessage(new AuthError("popup_window_error", "")), /Allow popups/);
  assert.doesNotMatch(signInErrorMessage(new AuthError("sensitive response value", "")), /sensitive response value/);
  assert.match(signInErrorMessage(new AuthError("sensitive response value", "")), /unknown_error/);
});