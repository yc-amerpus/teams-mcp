import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(server: McpServer, graphService: GraphService) {
  server.tool(
    "check_auth",
    "Check the authentication status of the Microsoft Graph connection. Verifies that app credentials are valid and the server can connect to Microsoft Graph.",
    {},
    async () => {
      const status = await graphService.getAuthStatus();
      return {
        content: [
          {
            type: "text",
            text: status.isAuthenticated
              ? `Authenticated. Tenant: ${status.tenantId}, Client: ${status.clientId}`
              : "Not authenticated. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables.",
          },
        ],
      };
    }
  );

  server.tool(
    "get_app_info",
    "Get information about the authenticated application, including tenant ID, client ID, and available permissions. Useful for debugging and understanding what the server can do.",
    {},
    async () => {
      const tenantId = process.env.AZURE_TENANT_ID || "not set";
      const clientId = process.env.AZURE_CLIENT_ID || "not set";

      const appInfo = {
        tenantId,
        clientId,
        authMode: process.env.AUTH_TOKEN ? "static_token" : "client_credentials",
        availablePermissions: [
          "Team.ReadBasic.All",
          "Channel.ReadBasic.All",
          "ChannelMessage.Read.All",
          "ChannelMessage.Send",
          "TeamMember.Read.All",
          "User.Read.All",
        ],
        capabilities: [
          "List teams in tenant",
          "List channels",
          "Read channel messages (with author/timestamps)",
          "Send/reply to channel messages",
          "List team members",
          "Search/get user profiles",
        ],
        limitations: [
          "No chat access (requires protected API approval)",
          "No search API (delegated-only)",
          "No current user context (app-only)",
        ],
      };

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(appInfo, null, 2),
          },
        ],
      };
    }
  );
}
