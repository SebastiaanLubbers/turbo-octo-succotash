import { AuthError, InteractionRequiredAuthError, PublicClientApplication, type AccountInfo } from "@azure/msal-browser";
import { create } from "zustand";
import { AppError } from "./errors";

const scopes = ["User.Read", "Sites.ReadWrite.All"];
const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID as string | undefined;
const tenantId = (import.meta.env.VITE_ENTRA_TENANT_ID as string | undefined) || "organizations";
const redirectUri = new URL(`${import.meta.env.BASE_URL}auth-callback.html`, window.location.origin + "/").href;

const client = clientId ? new PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
  },
  cache: { cacheLocation: "sessionStorage" },
  system: { enableLegacyPolling: true },
}) : null;

interface AuthState {
  ready: boolean;
  account: AccountInfo | null;
  error: string | null;
  initialize: () => Promise<void>;
  signIn: () => Promise<void>;
  signOut: () => Promise<void>;
}

let initialization: Promise<void> | undefined;

export function signInErrorMessage(error: unknown): string {
  if (!(error instanceof AuthError)) return "Sign-in did not complete. Try again or check your browser's popup settings.";
  const aadCode = error.errorMessage.match(/\bAADSTS\d{4,8}\b/i)?.[0].toUpperCase();
  const code = /^[a-z0-9_]{1,60}$/i.test(error.errorCode) ? error.errorCode : "unknown_error";
  const detail = aadCode ? `${code}, ${aadCode}` : code;
  if (aadCode === "AADSTS50011") {
    return `The redirect URI does not match the app registration. Add ${redirectUri} as a Single-page application redirect URI in Microsoft Entra. (${detail})`;
  }
  if (code === "user_cancelled" || code === "popup_window_error" || code === "empty_window_error") {
    return `Sign-in was cancelled or the popup could not open. Allow popups and try again. (${detail})`;
  }
  if (code === "timed_out" || code === "monitor_window_timeout") {
    return `Sign-in timed out. Check that the registered redirect URI is ${redirectUri} and that the callback page is reachable. (${detail})`;
  }
  return `Sign-in did not complete. Check the Microsoft Entra app registration and permissions. Error: ${detail}.`;
}

export const useBrowserAuth = create<AuthState>((set) => ({
  ready: false,
  account: null,
  error: null,
  initialize: async () => {
    if (!client) {
      set({ ready: true, error: "Microsoft sign-in is not configured. Set VITE_ENTRA_CLIENT_ID for this deployment." });
      return;
    }
    try {
      initialization ??= client.initialize();
      await initialization;
      const account = client.getActiveAccount() ?? client.getAllAccounts()[0] ?? null;
      client.setActiveAccount(account);
      set({ ready: true, account, error: null });
    } catch {
      set({ ready: true, error: "Microsoft sign-in could not be initialized." });
    }
  },
  signIn: async () => {
    if (!client) return;
    set({ error: null });
    try {
      const result = await client.loginPopup({ scopes });
      client.setActiveAccount(result.account);
      set({ account: result.account });
    } catch (error) {
      set({ error: signInErrorMessage(error) });
    }
  },
  signOut: async () => {
    if (!client) return;
    try {
      await client.clearCache();
      client.setActiveAccount(null);
      set({ account: null, error: null });
    } catch {
      set({ error: "Could not clear the Microsoft sign-in session. Try signing out again." });
    }
  },
}));

export async function graphAccessToken(): Promise<string> {
  const account = client?.getActiveAccount();
  if (!client || !account) throw new AppError("auth", "Sign in to Microsoft 365 before accessing SharePoint.", "Microsoft account unavailable", 401);
  try {
    return (await client.acquireTokenSilent({ account, scopes })).accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      throw new AppError("auth", "Your Microsoft 365 session needs attention. Sign out and sign in again.", "Microsoft authentication requires interaction", 401);
    }
    throw error;
  }
}