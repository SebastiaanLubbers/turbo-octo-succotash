// Error handling and result mapping.
// Every failure becomes an AppError with two separate messages:
//   userMessage — business-friendly text shown in the UI;
//   diagnostic  — short developer detail (operation + HTTP status), safe to show
//                 under "Details for support". Raw connector payloads never reach
//                 the UI; they are logged to the developer console only.

export type ErrorKind =
  | "auth"
  | "forbidden"
  | "not-found"
  | "conflict"
  | "throttled"
  | "policy"
  | "validation"
  | "service"
  | "unknown";

export class AppError extends Error {
  readonly kind: ErrorKind;
  readonly userMessage: string;
  readonly diagnostic: string;
  readonly status?: number;
  constructor(
    kind: ErrorKind,
    userMessage: string,
    diagnostic: string,
    status?: number,
  ) {
    super(userMessage);
    this.name = "AppError";
    this.kind = kind;
    this.userMessage = userMessage;
    this.diagnostic = diagnostic;
    this.status = status;
  }
}

export function statusOf(error: unknown): number | undefined {
  if (error && typeof error === "object" && "status" in error) {
    const s = (error as { status?: unknown }).status;
    if (typeof s === "number") return s;
  }
  return undefined;
}

export interface ErrorContext {
  /** Connector operation, e.g. "CreateNewFolder". */
  operation: string;
  /** What the user was doing, e.g. "open the site". Used in the user message. */
  action: string;
}

export function toAppError(error: unknown, ctx: ErrorContext): AppError {
  if (error instanceof AppError) return error;
  const status = statusOf(error);
  const diagnostic = `${ctx.operation} failed${status ? ` (HTTP ${status})` : ""} at ${new Date().toISOString()}`;
  // Developer console only — never rendered.
  console.error(`[SharePoint] ${ctx.operation}`, error);
  switch (status) {
    case 401:
      return new AppError(
        "auth",
        `Your Microsoft 365 sign-in is unavailable or has expired, so the app couldn't ${ctx.action}. Reload the app to sign in again — your entries are kept.`,
        diagnostic,
        status,
      );
    case 403:
      return new AppError(
        "forbidden",
        `You don't have permission to ${ctx.action}. Ask the site owner for edit (contribute) access.`,
        diagnostic,
        status,
      );
    case 404:
      return new AppError(
        "not-found",
        `The SharePoint location wasn't found or isn't accessible with your account, so the app couldn't ${ctx.action}.`,
        diagnostic,
        status,
      );
    case 409:
      return new AppError(
        "conflict",
        `An item with the same name already exists, so the app couldn't ${ctx.action}.`,
        diagnostic,
        status,
      );
    case 429:
      return new AppError(
        "throttled",
        `SharePoint is busy right now. Wait a moment and try again to ${ctx.action}.`,
        diagnostic,
        status,
      );
    default:
      if (status && status >= 500) {
        return new AppError(
          "service",
          `SharePoint had a temporary problem, so the app couldn't ${ctx.action}. Try again.`,
          diagnostic,
          status,
        );
      }
      return new AppError(
        "unknown",
        `The app couldn't ${ctx.action}. Try again; if it keeps failing, share the support details with your administrator.`,
        diagnostic,
        status,
      );
  }
}
