/**
 * Add-in configuration. Backend URL is overridable at runtime via
 * localStorage("ledger.backendUrl") so devs can point a sideloaded add-in at
 * any backend without rebuilding.
 *
 * Note (PLATFORM_QUIRKS): http://localhost is a trusted origin for mixed
 * content in Chromium (Excel web, Windows WebView2), but Safari/WKWebView on
 * Mac blocks https->http fetches even to localhost — run the backend behind
 * https or a tunnel when developing on Mac desktop.
 */

const DEFAULT_BACKEND_URL = "http://localhost:8000";

export function getBackendUrl(): string {
  try {
    return localStorage.getItem("ledger.backendUrl") ?? DEFAULT_BACKEND_URL;
  } catch {
    return DEFAULT_BACKEND_URL;
  }
}

export const APP_NAME = "Ledger";
export const PHASE = "Phase 0 — skeleton";

/**
 * Auth configuration, resolved at build time from the environment.
 *
 * `isProductionBuild` exists so `createAuthProvider` can refuse to fall back
 * to unauthenticated requests when the registration IDs are missing. A dev
 * build without them is a local convenience; a production build without them
 * is a bug that would ship an add-in talking to a real API anonymously.
 *
 * webpack's DefinePlugin substitutes these; the `typeof` guards keep the
 * module importable under vitest, where `process` exists but the plugin has
 * not run.
 */
export interface AuthConfig {
  mode: "naa" | "dev";
  clientId: string;
  authority: string;
  scopes: string[];
  isProductionBuild: boolean;
}

function env(name: string): string {
  try {
    return (typeof process !== "undefined" && process.env?.[name]) || "";
  } catch {
    return "";
  }
}

export function getAuthConfig(): AuthConfig {
  const clientId = env("LEDGER_ENTRA_CLIENT_ID");
  const tenantId = env("LEDGER_ENTRA_TENANT_ID");
  const apiClientId = env("LEDGER_ENTRA_API_CLIENT_ID") || clientId;
  const isProductionBuild = env("NODE_ENV") === "production";

  if (!clientId || !tenantId) {
    return { mode: "dev", clientId: "", authority: "", scopes: [], isProductionBuild };
  }
  return {
    mode: "naa",
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    // The API's own registration is the audience the backend validates.
    scopes: [`api://${apiClientId}/access`],
    isProductionBuild,
  };
}
