import { useEffect, useState } from "react";
import { NavLink, Outlet, useLocation, useNavigate } from "react-router-dom";
import { FolderTree, LogIn, LogOut, Settings2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Toaster } from "@/components/ui/sonner";
import { routes } from "@/routes";
import { useWizard } from "@/state/wizard-store";
import { cn } from "@/lib/utils";
import { queryClient } from "@/lib/query-client";
import { useBrowserAuth } from "@/services/browser-auth";

// App frame: product identity, page navigation and the global "change configuration" control.
export function AppShell() {
  const config = useWizard((s) => s.config);
  const busy = useWizard((s) => s.busy);
  const step = useWizard((s) => s.step);
  const startReconfigure = useWizard((s) => s.startReconfigure);
  const { pathname } = useLocation();
  const navigate = useNavigate();
  const { ready, account, error, initialize, signIn, signOut } = useBrowserAuth();
  const [authBusy, setAuthBusy] = useState(false);
  const showReconfigure = Boolean(config) && (pathname !== "/" || step >= 6);
  useEffect(() => { void initialize(); }, [initialize]);

  if (!ready || !account) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-background px-4">
        <main className="w-full max-w-md space-y-5 rounded-lg border bg-card p-6">
          <div className="flex items-center gap-3">
            <span className="flex size-10 items-center justify-center rounded-md bg-primary text-primary-foreground"><FolderTree className="size-5" aria-hidden="true" /></span>
            <h1 className="text-lg font-semibold">SharePoint Structure Builder</h1>
          </div>
          {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
          <Button disabled={!ready || authBusy || Boolean(error?.includes("not configured"))} onClick={() => {
            setAuthBusy(true);
            void signIn().finally(() => setAuthBusy(false));
          }}><LogIn aria-hidden="true" /> {authBusy ? "Signing in..." : "Sign in with Microsoft"}</Button>
        </main>
        <Toaster />
      </div>
    );
  }
  return (
    <div className="flex min-h-screen flex-col bg-background">
      <a
        href="#wizard-content"
        className="sr-only z-50 rounded-md bg-primary px-3 py-2 text-primary-foreground focus:not-sr-only focus:fixed focus:top-2 focus:left-2"
      >
        Skip to main content
      </a>
      <header className="sticky top-0 z-40 border-b bg-card">
        <div className="mx-auto flex h-14 w-full max-w-[1400px] items-center gap-3 px-4 md:px-6">
          <span className="flex size-8 shrink-0 items-center justify-center rounded-md bg-primary text-primary-foreground" aria-hidden="true">
            <FolderTree className="size-4.5" />
          </span>
          <span className="hidden font-semibold sm:inline">SharePoint Structure Builder</span>
          <nav aria-label="Main" className="ml-2 flex items-center gap-1 sm:ml-6">
            {routes
              .filter((r) => r.label)
              .map((r) => {
                const Icon = r.icon;
                return (
                  <NavLink
                    key={r.path}
                    to={r.path}
                    end
                    onClick={(e) => {
                      if (busy) e.preventDefault();
                    }}
                    className={({ isActive }) =>
                      cn(
                        "inline-flex items-center gap-1.5 rounded-md px-2.5 py-1.5 text-sm hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring",
                        isActive && "bg-accent font-medium",
                      )
                    }
                  >
                    {Icon && <Icon className="size-4" aria-hidden="true" />}
                    {r.label}
                  </NavLink>
                );
              })}
          </nav>
          {error && <span role="alert" className="text-sm text-destructive">{error}</span>}
          {showReconfigure && (
            <Button
              variant="ghost"
              size="sm"
              className="ml-auto"
              disabled={busy}
              onClick={() => {
                startReconfigure();
                navigate("/");
              }}
            >
              <Settings2 aria-hidden="true" /> <span className="hidden md:inline">Change configuration</span>
              <span className="sr-only md:hidden">Change configuration</span>
            </Button>
          )}
          <Button variant="ghost" size="sm" className={showReconfigure ? "" : "ml-auto"} disabled={busy || authBusy} onClick={() => {
            setAuthBusy(true);
            void signOut().then(() => queryClient.clear()).finally(() => setAuthBusy(false));
          }} title="Sign out of this app"><LogOut aria-hidden="true" /><span className="hidden md:inline">Sign out</span><span className="sr-only md:hidden">Sign out</span></Button>
        </div>
      </header>
      <main className="flex flex-1 flex-col">
        <Outlet />
      </main>
      <Toaster />
    </div>
  );
}

export default AppShell;
