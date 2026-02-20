import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock external dependencies
vi.mock("@modelcontextprotocol/sdk/server/mcp.js");
vi.mock("@modelcontextprotocol/sdk/server/stdio.js");

// Mock console methods
const mockConsoleError = vi.fn();
const mockProcessExit = vi.fn();

// Setup global mocks
beforeEach(() => {
  vi.clearAllMocks();

  // Mock console methods
  vi.spyOn(console, "error").mockImplementation(mockConsoleError);
  vi.spyOn(process, "exit").mockImplementation(mockProcessExit as any);

  // Reset process.argv
  process.argv = ["node", "index.js"];
});

describe("MCP Server Integration", () => {
  describe("Environment Variable Validation", () => {
    it("should exit with error when AZURE_TENANT_ID is missing", async () => {
      const originalTenant = process.env.AZURE_TENANT_ID;
      const originalClient = process.env.AZURE_CLIENT_ID;
      const originalSecret = process.env.AZURE_CLIENT_SECRET;
      const originalToken = process.env.AUTH_TOKEN;

      delete process.env.AZURE_TENANT_ID;
      delete process.env.AZURE_CLIENT_ID;
      delete process.env.AZURE_CLIENT_SECRET;
      delete process.env.AUTH_TOKEN;

      await import("../index.js");

      expect(mockConsoleError).toHaveBeenCalledWith(
        expect.stringContaining("Missing required environment variables")
      );
      expect(mockProcessExit).toHaveBeenCalledWith(1);

      // Restore
      if (originalTenant !== undefined) process.env.AZURE_TENANT_ID = originalTenant;
      if (originalClient !== undefined) process.env.AZURE_CLIENT_ID = originalClient;
      if (originalSecret !== undefined) process.env.AZURE_CLIENT_SECRET = originalSecret;
      if (originalToken !== undefined) process.env.AUTH_TOKEN = originalToken;
    });

    it("should skip env var validation when AUTH_TOKEN is set", async () => {
      const originalTenant = process.env.AZURE_TENANT_ID;
      const originalClient = process.env.AZURE_CLIENT_ID;
      const originalSecret = process.env.AZURE_CLIENT_SECRET;

      delete process.env.AZURE_TENANT_ID;
      delete process.env.AZURE_CLIENT_ID;
      delete process.env.AZURE_CLIENT_SECRET;
      process.env.AUTH_TOKEN = "test-token";

      const { McpServer } = await import("@modelcontextprotocol/sdk/server/mcp.js");
      const { StdioServerTransport } = await import("@modelcontextprotocol/sdk/server/stdio.js");

      const mockServer = {
        tool: vi.fn(),
        connect: vi.fn(),
      };
      const mockTransport = {};

      vi.mocked(McpServer).mockImplementation(() => mockServer as any);
      vi.mocked(StdioServerTransport).mockImplementation(() => mockTransport as any);

      await import("../index.js");

      // Should NOT have called process.exit
      expect(mockProcessExit).not.toHaveBeenCalled();

      // Restore
      delete process.env.AUTH_TOKEN;
      if (originalTenant !== undefined) process.env.AZURE_TENANT_ID = originalTenant;
      if (originalClient !== undefined) process.env.AZURE_CLIENT_ID = originalClient;
      if (originalSecret !== undefined) process.env.AZURE_CLIENT_SECRET = originalSecret;
    });
  });
});
