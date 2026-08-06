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
