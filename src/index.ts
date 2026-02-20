#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { GraphService } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";

function validateEnv(): void {
  // Skip validation when using direct token injection
  if (process.env.AUTH_TOKEN) return;

  const required = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"];
  const missing = required.filter((key) => !process.env[key]);
  if (missing.length > 0) {
    console.error(`Missing required environment variables: ${missing.join(", ")}`);
    console.error("Set these in your environment or .env file.");
    process.exit(1);
  }
}

async function startMcpServer() {
  const server = new McpServer({
    name: "teams-mcp",
    version: "0.7.0",
  });

  const graphService = GraphService.getInstance();

  registerAuthTools(server, graphService);
  registerUsersTools(server, graphService);
  registerTeamsTools(server, graphService);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Graph MCP Server started (app-only mode)");
}

async function main() {
  validateEnv();
  await startMcpServer();
}

// Handle uncaught errors
process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});
