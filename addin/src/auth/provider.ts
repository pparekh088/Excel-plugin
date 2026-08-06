/**
 * Auth provider (P1-6).
 *
 * Two implementations behind one interface:
 *
 *   NaaAuthProvider  Nested App Authentication. MSAL's
 *                    `createNestablePublicClientApplication` lets an add-in
 *                    running inside Office acquire tokens through the HOST's
 *                    identity broker — no popup, no separate sign-in, and the
 *                    token is scoped to the Ledger API registration.
 *   DevAuthProvider  returns null, for a backend running LEDGER_AUTH_MODE=dev.
 *
 * The selection is not a preference. `createAuthProvider` throws rather than
 * hand back the dev provider when the build is configured for production: a
 * silent fallback to "no token" is how an add-in ends up talking to a
 * production API anonymously, and the failure would look like a working
 * system right up until it did not.
 *
 * Tokens are cached in memory with a skew margin. MSAL caches too, but the
 * custom-functions runtime calls this on every batch and an await per batch
 * that resolves from our own map is cheaper than one that crosses into MSAL.
 */

import { getAuthConfig } from "../config";

export interface AuthProvider {
  /** A bearer token, or null when the backend accepts anonymous (dev only). */
  getAccessToken(): Promise<string | null>;
}

export class DevAuthProvider implements AuthProvider {
  async getAccessToken(): Promise<string | null> {
    return null;
  }
}

/** Refresh this long before expiry rather than racing the clock. */
const EXPIRY_SKEW_MS = 5 * 60 * 1000;

interface CachedToken {
  token: string;
  expiresAt: number;
}

/**
 * Minimal shape of what MSAL gives us. Declared rather than imported at the
 * top level so the custom-functions bundle does not pull MSAL in until a token
 * is actually needed — that runtime is loaded on every recalculation.
 */
interface NestablePublicClientApplication {
  acquireTokenSilent(request: {
    scopes: string[];
    account?: unknown;
  }): Promise<{ accessToken: string; expiresOn: Date | null; account: unknown }>;
  acquireTokenPopup(request: {
    scopes: string[];
  }): Promise<{ accessToken: string; expiresOn: Date | null; account: unknown }>;
  getAllAccounts(): unknown[];
}

export class NaaAuthProvider implements AuthProvider {
  private client: Promise<NestablePublicClientApplication> | null = null;
  private cached: CachedToken | null = null;

  constructor(
    private readonly clientId: string,
    private readonly authority: string,
    private readonly scopes: string[]
  ) {}

  private async getClient(): Promise<NestablePublicClientApplication> {
    if (!this.client) {
      this.client = import("@azure/msal-browser").then((msal) =>
        msal.createNestablePublicClientApplication({
          auth: { clientId: this.clientId, authority: this.authority },
          // Office hosts the add-in in an iframe; sessionStorage keeps the
          // cache scoped to this task pane instance.
          cache: { cacheLocation: "sessionStorage" },
        })
      ) as Promise<NestablePublicClientApplication>;
    }
    return this.client;
  }

  async getAccessToken(): Promise<string> {
    const now = Date.now();
    if (this.cached && this.cached.expiresAt - EXPIRY_SKEW_MS > now) {
      return this.cached.token;
    }

    const client = await this.getClient();
    const accounts = client.getAllAccounts();
    let result;
    try {
      result = await client.acquireTokenSilent({
        scopes: this.scopes,
        ...(accounts.length > 0 ? { account: accounts[0] } : {}),
      });
    } catch {
      // Silent acquisition fails on first use and when consent is needed.
      // Under NAA the host brokers this without a separate sign-in window.
      result = await client.acquireTokenPopup({ scopes: this.scopes });
    }

    this.cached = {
      token: result.accessToken,
      expiresAt: result.expiresOn ? result.expiresOn.getTime() : now + 30 * 60 * 1000,
    };
    return result.accessToken;
  }
}

/**
 * One provider per runtime. The task pane and the custom-functions runtime are
 * separate JavaScript contexts unless the manifest declares a shared runtime;
 * with one declared they share this module instance and therefore the cache.
 */
let singleton: AuthProvider | null = null;

export function createAuthProvider(): AuthProvider {
  if (singleton) return singleton;
  const config = getAuthConfig();

  if (config.mode === "naa") {
    singleton = new NaaAuthProvider(config.clientId, config.authority, config.scopes);
    return singleton;
  }

  if (config.isProductionBuild) {
    // Never silently downgrade to anonymous in a production build.
    throw new Error(
      "Auth is not configured (LEDGER_ENTRA_CLIENT_ID / LEDGER_ENTRA_TENANT_ID are unset) " +
        "and this is a production build. Refusing to fall back to unauthenticated requests."
    );
  }

  singleton = new DevAuthProvider();
  return singleton;
}

/** Tests only: drop the cached provider. */
export function resetAuthProvider(): void {
  singleton = null;
}
