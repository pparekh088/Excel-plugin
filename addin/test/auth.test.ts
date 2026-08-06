/**
 * Auth wiring (P1-6).
 *
 * Two things are being pinned here, and the second matters more than the first:
 *
 *  1. AI.* custom functions send a bearer token. They used to send none at
 *     all, which works perfectly against a dev-mode backend and fails every
 *     cell in every workbook the moment the backend requires Entra.
 *  2. A production build with no registration configured REFUSES to run rather
 *     than falling back to anonymous requests. A silent downgrade is how an
 *     add-in ends up talking to a real API with no identity, and it looks like
 *     a working system until it does not.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const ORIGINAL_ENV = { ...process.env };

function setEnv(values: Record<string, string | undefined>): void {
  for (const [key, value] of Object.entries(values)) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
}

beforeEach(() => {
  vi.resetModules();
  setEnv({
    NODE_ENV: "development",
    LEDGER_ENTRA_CLIENT_ID: undefined,
    LEDGER_ENTRA_TENANT_ID: undefined,
    LEDGER_ENTRA_API_CLIENT_ID: undefined,
  });
});

afterEach(() => {
  process.env = { ...ORIGINAL_ENV };
});

describe("getAuthConfig", () => {
  it("falls back to dev mode when no registration is configured", async () => {
    const { getAuthConfig } = await import("../src/config");
    const config = getAuthConfig();
    expect(config.mode).toBe("dev");
    expect(config.isProductionBuild).toBe(false);
  });

  it("builds the NAA config from the registration IDs", async () => {
    setEnv({
      LEDGER_ENTRA_CLIENT_ID: "addin-client-id",
      LEDGER_ENTRA_TENANT_ID: "tenant-guid",
      LEDGER_ENTRA_API_CLIENT_ID: "api-client-id",
    });
    const { getAuthConfig } = await import("../src/config");
    const config = getAuthConfig();

    expect(config.mode).toBe("naa");
    expect(config.clientId).toBe("addin-client-id");
    expect(config.authority).toBe("https://login.microsoftonline.com/tenant-guid");
    // The audience the backend validates is the API registration, not the
    // add-in's own — getting this wrong yields a token the server rejects.
    expect(config.scopes).toEqual(["api://api-client-id/access"]);
  });

  it("defaults the API registration to the add-in's when only one is given", async () => {
    setEnv({ LEDGER_ENTRA_CLIENT_ID: "one-id", LEDGER_ENTRA_TENANT_ID: "tenant" });
    const { getAuthConfig } = await import("../src/config");
    expect(getAuthConfig().scopes).toEqual(["api://one-id/access"]);
  });
});

describe("createAuthProvider", () => {
  it("returns the dev provider for a dev build with no registration", async () => {
    const { createAuthProvider } = await import("../src/auth/provider");
    expect(await createAuthProvider().getAccessToken()).toBeNull();
  });

  it("REFUSES to fall back to anonymous in a production build", async () => {
    setEnv({ NODE_ENV: "production" });
    const { createAuthProvider } = await import("../src/auth/provider");
    expect(() => createAuthProvider()).toThrow(/Refusing to fall back/);
  });

  it("uses NAA once the registration is configured", async () => {
    setEnv({
      NODE_ENV: "production",
      LEDGER_ENTRA_CLIENT_ID: "addin-client-id",
      LEDGER_ENTRA_TENANT_ID: "tenant-guid",
    });
    const { createAuthProvider, NaaAuthProvider } = await import("../src/auth/provider");
    expect(createAuthProvider()).toBeInstanceOf(NaaAuthProvider);
  });

  it("hands every caller the same instance, so the token cache is shared", async () => {
    const { createAuthProvider } = await import("../src/auth/provider");
    expect(createAuthProvider()).toBe(createAuthProvider());
  });
});

describe("NaaAuthProvider token caching", () => {
  /** Stand in for MSAL without reaching the network. */
  function providerWith(acquire: () => Promise<unknown>, NaaAuthProvider: unknown) {
    const Ctor = NaaAuthProvider as new (
      clientId: string,
      authority: string,
      scopes: string[]
    ) => { getAccessToken(): Promise<string> };
    const provider = Ctor.prototype as unknown as Record<string, unknown>;
    const instance = new Ctor("id", "https://login.microsoftonline.com/t", ["api://x/access"]);
    (instance as unknown as Record<string, unknown>).client = Promise.resolve({
      getAllAccounts: () => [{ homeAccountId: "acct" }],
      acquireTokenSilent: acquire,
      acquireTokenPopup: acquire,
    });
    void provider;
    return instance;
  }

  it("reuses a token until it is close to expiry", async () => {
    const { NaaAuthProvider } = await import("../src/auth/provider");
    const acquire = vi.fn(async () => ({
      accessToken: "tok-1",
      expiresOn: new Date(Date.now() + 60 * 60 * 1000),
      account: {},
    }));
    const provider = providerWith(acquire, NaaAuthProvider);

    expect(await provider.getAccessToken()).toBe("tok-1");
    expect(await provider.getAccessToken()).toBe("tok-1");
    // One acquisition, not one per call — a recalculation issues many batches.
    expect(acquire).toHaveBeenCalledTimes(1);
  });

  it("re-acquires when the cached token is inside the skew window", async () => {
    const { NaaAuthProvider } = await import("../src/auth/provider");
    let counter = 0;
    const acquire = vi.fn(async () => ({
      accessToken: `tok-${++counter}`,
      // Expires in one minute: inside the five-minute skew, so never reused.
      expiresOn: new Date(Date.now() + 60 * 1000),
      account: {},
    }));
    const provider = providerWith(acquire, NaaAuthProvider);

    expect(await provider.getAccessToken()).toBe("tok-1");
    expect(await provider.getAccessToken()).toBe("tok-2");
  });

  it("falls back to the brokered popup when silent acquisition fails", async () => {
    const { NaaAuthProvider } = await import("../src/auth/provider");
    const silent = vi.fn(async () => {
      throw new Error("interaction_required");
    });
    const popup = vi.fn(async () => ({
      accessToken: "tok-popup",
      expiresOn: new Date(Date.now() + 60 * 60 * 1000),
      account: {},
    }));

    const Ctor = NaaAuthProvider;
    const instance = new Ctor("id", "https://login.microsoftonline.com/t", ["api://x/access"]);
    (instance as unknown as Record<string, unknown>).client = Promise.resolve({
      getAllAccounts: () => [],
      acquireTokenSilent: silent,
      acquireTokenPopup: popup,
    });

    expect(await instance.getAccessToken()).toBe("tok-popup");
    expect(silent).toHaveBeenCalled();
    expect(popup).toHaveBeenCalled();
  });
});

describe("AI.* custom functions send the token", () => {
  beforeEach(() => {
    (globalThis as Record<string, unknown>).CustomFunctions = {
      associate: vi.fn(),
    };
  });

  it("attaches an Authorization header when a token is available", async () => {
    setEnv({
      LEDGER_ENTRA_CLIENT_ID: "addin-client-id",
      LEDGER_ENTRA_TENANT_ID: "tenant-guid",
    });
    const provider = await import("../src/auth/provider");
    vi.spyOn(provider, "createAuthProvider").mockReturnValue({
      getAccessToken: async () => "the-token",
    });

    const fetchMock = vi.fn(async (_url: string, _init: RequestInit) => ({
      ok: true,
      status: 200,
      // The server's real envelope shape: {ok, value, error} per request.
      json: async () => ({ results: [{ ok: true, value: "yes" }] }),
    }));
    (globalThis as Record<string, unknown>).fetch = fetchMock;

    const functions = await import("../src/functions/functions");
    await functions.aiClassify("some text", "yes,no");

    expect(fetchMock).toHaveBeenCalled();
    const init = fetchMock.mock.calls[0]![1] as unknown as {
      headers: Record<string, string>;
    };
    expect(init.headers.Authorization).toBe("Bearer the-token");
  });

  it("sends no Authorization header in dev, where the backend accepts anonymous", async () => {
    const fetchMock = vi.fn(async (_url: string, _init: RequestInit) => ({
      ok: true,
      status: 200,
      // The server's real envelope shape: {ok, value, error} per request.
      json: async () => ({ results: [{ ok: true, value: "yes" }] }),
    }));
    (globalThis as Record<string, unknown>).fetch = fetchMock;

    const functions = await import("../src/functions/functions");
    await functions.aiClassify("some text", "yes,no");

    const init = fetchMock.mock.calls[0]![1] as unknown as {
      headers: Record<string, string>;
    };
    expect(init.headers.Authorization).toBeUndefined();
  });

  it("says 'sign in' rather than 'the model failed' on a 401", async () => {
    (globalThis as Record<string, unknown>).fetch = vi.fn(async () => ({
      ok: false,
      status: 401,
      json: async () => ({}),
    }));

    const functions = await import("../src/functions/functions");
    // The error propagates to Excel, which shows it on the cell. What matters
    // is that the message names the cause the user can act on rather than
    // reading as a model failure.
    await expect(functions.aiClassify("some text", "yes,no")).rejects.toThrow(/sign in/);
  });
});
