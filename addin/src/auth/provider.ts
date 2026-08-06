/**
 * Auth provider seam.
 *
 * Phase 0 ships DevAuthProvider (no token; backend runs LEDGER_AUTH_MODE=dev).
 * The Entra path is Nested App Authentication (NAA) via MSAL.js
 * `createNestablePublicClientApplication`, reusing the org's existing Entra
 * app-registration pattern: the add-in SPA registration acquires an access
 * token for the Ledger API registration (scope api://<api-client-id>/access),
 * and the backend validates it against the tenant JWKS (see server auth.py).
 * Wiring NAA requires the real registration IDs and the manifest
 * WebApplicationInfo block — tracked for the pilot deployment, not needed for
 * the Phase 0 gate.
 */

export interface AuthProvider {
  /** Returns a bearer token, or null when the backend accepts anonymous (dev). */
  getAccessToken(): Promise<string | null>;
}

export class DevAuthProvider implements AuthProvider {
  async getAccessToken(): Promise<string | null> {
    return null;
  }
}

export function createAuthProvider(): AuthProvider {
  return new DevAuthProvider();
}
