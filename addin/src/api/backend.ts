/**
 * The one place add-in code talks to the Ledger backend (R2 items 2/3/7).
 *
 * There used to be three fetch call sites — task pane client, AI batch,
 * AI.FORECAST — and they drifted: batch gained the bearer token when auth
 * landed, forecast kept posting anonymously, so under an Entra backend four of
 * the five AI.* functions authenticated and the fifth returned 401. A second
 * fetch implementation is a second chance to forget a header; this module
 * exists so there is nothing to forget.
 *
 * It also carries the active session ID. The server meters AI spend per
 * (principal, session); a batch call without the session lands in the
 * principal's catch-all bucket, mixing every open workbook's spend together.
 * The task pane stores the server-issued ID here when it connects, and the
 * custom functions — same module instance under the shared runtime — pick it
 * up on every batch. localStorage mirrors it for hosts without a shared
 * runtime, where the two contexts do not share module state.
 */

import { createAuthProvider } from "../auth/provider";
import { getBackendUrl } from "../config";

const SESSION_KEY = "ledger.sessionId";

let activeSessionId: string | null = null;

/** Called by the task pane when the backend issues a session. */
export function setActiveSessionId(sessionId: string | null): void {
  activeSessionId = sessionId;
  try {
    if (sessionId === null) localStorage.removeItem(SESSION_KEY);
    else localStorage.setItem(SESSION_KEY, sessionId);
  } catch {
    /* no localStorage in this runtime; the module variable still works */
  }
}

export function getActiveSessionId(): string | null {
  if (activeSessionId !== null) return activeSessionId;
  try {
    return localStorage.getItem(SESSION_KEY);
  } catch {
    return null;
  }
}

/**
 * POST JSON to the backend with the bearer token attached (when auth is
 * configured — dev mode sends none and the dev backend accepts that).
 *
 * 401/403 are turned into a message that names the one cause the user can act
 * on. "The model failed" and "you are not signed in" demand different
 * responses, and only one of them is the user's to make.
 */
export async function postToBackend<T>(path: string, body: unknown): Promise<T> {
  const token = await createAuthProvider().getAccessToken();
  const response = await fetch(`${getBackendUrl()}${path}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      throw new Error(
        `Backend rejected the request (${response.status}) — sign in from the Ledger ` +
          `task pane, then recalculate.`
      );
    }
    throw new Error(`Backend returned ${response.status} for ${path}`);
  }
  return (await response.json()) as T;
}
