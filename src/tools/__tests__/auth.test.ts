import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  createMockGraphService,
  createMockMcpServer,
  createMockUnauthenticatedGraphService,
} from "../../test-utils/setup.js";
import { registerAuthTools } from "../auth.js";

describe("Authentication Tools", () => {
  let mockServer: any;
  let mockGraphService: any;

  beforeEach(() => {
    mockServer = createMockMcpServer();
    vi.clearAllMocks();
  });

  describe("check_auth tool", () => {
    it("should register check_auth tool correctly", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      expect(mockServer.tool).toHaveBeenCalledWith(
        "check_auth",
        expect.stringContaining("Check the authentication status"),
        {},
        expect.any(Function)
      );
    });

    it("should return authenticated status with tenant and client info", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const authTool = mockServer.getTool("check_auth");
      const result = await authTool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "Authenticated. Tenant: test-tenant-id, Client: test-client-id",
          },
        ],
      });

      expect(mockGraphService.getAuthStatus).toHaveBeenCalledTimes(1);
    });

    it("should return unauthenticated status when credentials are invalid", async () => {
      mockGraphService = createMockUnauthenticatedGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const authTool = mockServer.getTool("check_auth");
      const result = await authTool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "Not authenticated. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables.",
          },
        ],
      });

      expect(mockGraphService.getAuthStatus).toHaveBeenCalledTimes(1);
    });

    it("should handle authentication status errors", async () => {
      const errorMockGraphService = {
        getAuthStatus: vi.fn().mockRejectedValue(new Error("Auth check failed")),
      } as any;

      registerAuthTools(mockServer, errorMockGraphService);

      const authTool = mockServer.getTool("check_auth");

      await expect(authTool.handler()).rejects.toThrow("Auth check failed");
    });
  });

  describe("get_app_info tool", () => {
    it("should register get_app_info tool correctly", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      expect(mockServer.tool).toHaveBeenCalledWith(
        "get_app_info",
        expect.stringContaining("Get information about the authenticated application"),
        {},
        expect.any(Function)
      );
    });

    it("should return app info with env vars set", async () => {
      const originalTenant = process.env.AZURE_TENANT_ID;
      const originalClient = process.env.AZURE_CLIENT_ID;
      const originalToken = process.env.AUTH_TOKEN;

      process.env.AZURE_TENANT_ID = "my-tenant";
      process.env.AZURE_CLIENT_ID = "my-client";
      delete process.env.AUTH_TOKEN;

      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_app_info");
      const result = await tool.handler();
      const appInfo = JSON.parse(result.content[0].text);

      expect(appInfo.tenantId).toBe("my-tenant");
      expect(appInfo.clientId).toBe("my-client");
      expect(appInfo.authMode).toBe("client_credentials");
      expect(appInfo.availablePermissions).toContain("Team.ReadBasic.All");
      expect(appInfo.capabilities).toContain("List teams in tenant");
      expect(appInfo.limitations).toContain("No chat access (requires protected API approval)");

      // Restore
      if (originalTenant === undefined) delete process.env.AZURE_TENANT_ID;
      else process.env.AZURE_TENANT_ID = originalTenant;
      if (originalClient === undefined) delete process.env.AZURE_CLIENT_ID;
      else process.env.AZURE_CLIENT_ID = originalClient;
      if (originalToken === undefined) delete process.env.AUTH_TOKEN;
      else process.env.AUTH_TOKEN = originalToken;
    });

    it("should show static_token auth mode when AUTH_TOKEN is set", async () => {
      const originalToken = process.env.AUTH_TOKEN;
      process.env.AUTH_TOKEN = "some-token";

      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_app_info");
      const result = await tool.handler();
      const appInfo = JSON.parse(result.content[0].text);

      expect(appInfo.authMode).toBe("static_token");

      if (originalToken === undefined) delete process.env.AUTH_TOKEN;
      else process.env.AUTH_TOKEN = originalToken;
    });

    it("should show 'not set' when env vars are missing", async () => {
      const originalTenant = process.env.AZURE_TENANT_ID;
      const originalClient = process.env.AZURE_CLIENT_ID;

      delete process.env.AZURE_TENANT_ID;
      delete process.env.AZURE_CLIENT_ID;

      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_app_info");
      const result = await tool.handler();
      const appInfo = JSON.parse(result.content[0].text);

      expect(appInfo.tenantId).toBe("not set");
      expect(appInfo.clientId).toBe("not set");

      if (originalTenant === undefined) delete process.env.AZURE_TENANT_ID;
      else process.env.AZURE_TENANT_ID = originalTenant;
      if (originalClient === undefined) delete process.env.AZURE_CLIENT_ID;
      else process.env.AZURE_CLIENT_ID = originalClient;
    });
  });

  describe("tool registration", () => {
    it("should register all expected authentication tools", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService);

      const registeredTools = mockServer.getAllTools();
      expect(registeredTools).toContain("check_auth");
      expect(registeredTools).toContain("get_app_info");
      expect(registeredTools).toHaveLength(2);
    });

    it("should handle GraphService being undefined", () => {
      expect(() => {
        registerAuthTools(mockServer, undefined as any);
      }).not.toThrow();

      expect(mockServer.tool).toHaveBeenCalledTimes(2);
    });
  });

  describe("authentication state changes", () => {
    it("should reflect real-time authentication status changes", async () => {
      let isAuthenticated = false;
      const dynamicMockGraphService = {
        getAuthStatus: vi.fn().mockImplementation(() => {
          return Promise.resolve({
            isAuthenticated,
            tenantId: isAuthenticated ? "test-tenant" : undefined,
            clientId: isAuthenticated ? "test-client" : undefined,
          });
        }),
      } as any;

      registerAuthTools(mockServer, dynamicMockGraphService);
      const authTool = mockServer.getTool("check_auth");

      // Check unauthenticated status
      let result = await authTool.handler();
      expect(result.content[0].text).toContain("Not authenticated");

      // Simulate authentication
      isAuthenticated = true;

      // Check authenticated status
      result = await authTool.handler();
      expect(result.content[0].text).toContain("Authenticated");
      expect(result.content[0].text).toContain("test-tenant");
    });
  });
});
