import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { server } from "../../test-utils/setup.js";

// Mock @azure/msal-node
vi.mock("@azure/msal-node", () => ({
  ConfidentialClientApplication: vi.fn(),
}));

// Mock @microsoft/microsoft-graph-client
vi.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    initWithMiddleware: vi.fn(),
  },
}));

// Import after mocks are set up
import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { GraphService } from "../graph.js";

/** Set up the default MSAL mock: acquireTokenByClientCredential succeeds */
function setupDefaultMsalMock() {
  vi.mocked(ConfidentialClientApplication).mockImplementation(function () {
    return {
      acquireTokenByClientCredential: vi.fn().mockResolvedValue({
        accessToken: "mock-access-token",
        expiresOn: new Date(Date.now() + 3600000),
      }),
    };
  } as any);
}

describe("GraphService", () => {
  let graphService: GraphService;
  const originalEnv = { ...process.env };

  beforeEach(() => {
    server.listen({ onUnhandledRequest: "error" });
    vi.clearAllMocks();
    setupDefaultMsalMock();

    // Set required env vars
    process.env.AZURE_TENANT_ID = "test-tenant-id";
    process.env.AZURE_CLIENT_ID = "test-client-id";
    process.env.AZURE_CLIENT_SECRET = "test-client-secret";
    delete process.env.AUTH_TOKEN;

    // Reset GraphService singleton
    (GraphService as any).instance = undefined;
    graphService = GraphService.getInstance();
  });

  afterEach(() => {
    server.resetHandlers();
    server.close();
    // Restore env
    process.env.AZURE_TENANT_ID = originalEnv.AZURE_TENANT_ID;
    process.env.AZURE_CLIENT_ID = originalEnv.AZURE_CLIENT_ID;
    process.env.AZURE_CLIENT_SECRET = originalEnv.AZURE_CLIENT_SECRET;
    if (originalEnv.AUTH_TOKEN === undefined) {
      delete process.env.AUTH_TOKEN;
    } else {
      process.env.AUTH_TOKEN = originalEnv.AUTH_TOKEN;
    }
  });

  describe("getInstance", () => {
    it("should return singleton instance", () => {
      const instance1 = GraphService.getInstance();
      const instance2 = GraphService.getInstance();

      expect(instance1).toBe(instance2);
    });
  });

  describe("getAuthStatus", () => {
    it("should return unauthenticated when env vars are missing", async () => {
      delete process.env.AZURE_TENANT_ID;
      delete process.env.AZURE_CLIENT_ID;
      delete process.env.AZURE_CLIENT_SECRET;

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });

    it("should return unauthenticated when acquireTokenByClientCredential fails", async () => {
      vi.mocked(ConfidentialClientApplication).mockImplementationOnce(function () {
        return {
          acquireTokenByClientCredential: vi
            .fn()
            .mockRejectedValue(new Error("Invalid client secret")),
        };
      } as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });

    it("should return unauthenticated when acquireTokenByClientCredential returns null", async () => {
      vi.mocked(ConfidentialClientApplication).mockImplementationOnce(function () {
        return {
          acquireTokenByClientCredential: vi.fn().mockResolvedValue(null),
        };
      } as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });

    it("should return authenticated status with valid client credentials", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [{ id: "org-id", displayName: "Test Org" }] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({
        isAuthenticated: true,
        tenantId: "test-tenant-id",
        clientId: "test-client-id",
      });
    });

    it("should call /organization endpoint for auth verification", async () => {
      const mockGet = vi.fn().mockResolvedValue({ value: [{ id: "org-id" }] });
      const mockClient = {
        api: vi.fn().mockReturnValue({ get: mockGet }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      expect(mockClient.api).toHaveBeenCalledWith("/organization");
    });

    it("should handle Graph API errors gracefully", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error("API Error")),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status).toEqual({ isAuthenticated: false });
    });
  });

  describe("getClient", () => {
    it("should throw error when not authenticated", async () => {
      delete process.env.AZURE_TENANT_ID;

      await expect(graphService.getClient()).rejects.toThrow(
        "Not authenticated. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET"
      );
    });

    it("should return client when authenticated", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const client = await graphService.getClient();

      expect(client).toBeDefined();
    });
  });

  describe("isAuthenticated", () => {
    it("should return false when not initialized", () => {
      expect(graphService.isAuthenticated()).toBe(false);
    });

    it("should return true when client is initialized", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      expect(graphService.isAuthenticated()).toBe(true);
    });
  });

  describe("MSAL client credentials", () => {
    it("should create ConfidentialClientApplication with correct config", async () => {
      const mockAcquireToken = vi.fn().mockResolvedValue({
        accessToken: "mock-access-token",
        expiresOn: new Date(Date.now() + 3600000),
      });

      vi.mocked(ConfidentialClientApplication).mockImplementationOnce(function () {
        return {
          acquireTokenByClientCredential: mockAcquireToken,
        };
      } as any);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      expect(ConfidentialClientApplication).toHaveBeenCalledWith(
        expect.objectContaining({
          auth: expect.objectContaining({
            clientId: "test-client-id",
            clientSecret: "test-client-secret",
            authority: "https://login.microsoftonline.com/test-tenant-id",
          }),
        })
      );

      expect(mockAcquireToken).toHaveBeenCalledWith({
        scopes: ["https://graph.microsoft.com/.default"],
      });
    });

    it("should pass auth provider that calls acquireTokenByClientCredential", async () => {
      const mockAcquireToken = vi.fn().mockResolvedValue({
        accessToken: "mock-access-token",
        expiresOn: new Date(Date.now() + 3600000),
      });

      vi.mocked(ConfidentialClientApplication).mockImplementationOnce(function () {
        return {
          acquireTokenByClientCredential: mockAcquireToken,
        };
      } as any);

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getClient();

      // Extract the auth provider passed to Client.initWithMiddleware
      const initCall = vi.mocked(Client.initWithMiddleware).mock.calls[0];
      const authProvider = (initCall[0] as any).authProvider;

      // Call getAccessToken to verify it uses acquireTokenByClientCredential
      const token = await authProvider.getAccessToken();
      expect(token).toBe("mock-access-token");

      // acquireTokenByClientCredential should have been called (once during init + once via authProvider)
      expect(mockAcquireToken).toHaveBeenCalledTimes(2);
    });
  });

  describe("concurrent initialization", () => {
    it("should handle concurrent calls to getAuthStatus", async () => {
      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [{ id: "org-id" }] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const promises = [
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
        graphService.getAuthStatus(),
      ];

      const results = await Promise.all(promises);

      for (const result of results) {
        expect(result.isAuthenticated).toBe(true);
      }
    });
  });

  describe("AUTH_TOKEN environment variable", () => {
    it("should use AUTH_TOKEN from environment when provided", async () => {
      const mockPayload = btoa(JSON.stringify({ aud: "https://graph.microsoft.com" }));
      const validToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = validToken;

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [{ id: "org-id" }] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
      // MSAL should NOT be used when AUTH_TOKEN is set
      expect(ConfidentialClientApplication).not.toHaveBeenCalled();
    });

    it("should reject invalid JWT format from AUTH_TOKEN", async () => {
      process.env.AUTH_TOKEN = "invalid-token";

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(false);
    });

    it("should reject JWT without Graph audience from AUTH_TOKEN", async () => {
      const mockPayload = btoa(JSON.stringify({ aud: "https://other-service.com" }));
      const invalidToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = invalidToken;

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(false);
    });

    it("should handle JWT with audience as array from AUTH_TOKEN", async () => {
      const mockPayload = btoa(
        JSON.stringify({ aud: ["https://graph.microsoft.com", "https://other.com"] })
      );
      const validToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = validToken;

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [{ id: "org-id" }] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      const status = await graphService.getAuthStatus();

      expect(status.isAuthenticated).toBe(true);
    });

    it("should prefer AUTH_TOKEN over MSAL-based auth", async () => {
      const mockPayload = btoa(JSON.stringify({ aud: "https://graph.microsoft.com" }));
      const validToken = `header.${mockPayload}.signature`;
      process.env.AUTH_TOKEN = validToken;

      const mockClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ value: [{ id: "org-id" }] }),
        }),
      };

      vi.mocked(Client.initWithMiddleware).mockReturnValue(mockClient as any);

      await graphService.getAuthStatus();

      // MSAL should not be used when AUTH_TOKEN is present
      expect(ConfidentialClientApplication).not.toHaveBeenCalled();
    });
  });
});
