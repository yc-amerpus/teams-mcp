# Installation Guide

teams-mcp is an MCP server that provides Microsoft Teams integration via the Microsoft Graph API using app-only (client credentials) authentication. It supports two deployment modes:

- **stdio** — for VS Code, Claude Desktop, Claude Code, and other local MCP clients
- **SSE/HTTP** — for AnythingLLM, remote clients, or any HTTP-based MCP consumer (via supergateway)

## Prerequisites

### Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Azure Active Directory** > **App registrations** > **New registration**
2. Name it (e.g. `teams-mcp`), select **Single tenant**, click Register
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **Certificates & secrets** > **New client secret** > copy the secret value
5. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**, add:

| Permission | Purpose |
|---|---|
| `Team.ReadBasic.All` | List teams |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages with author/timestamps |
| `ChannelMessage.Send` | Send and reply to channel messages |
| `TeamMember.Read.All` | List team members |
| `User.Read.All` | Search and get user profiles |

6. Click **Grant admin consent** (requires tenant admin)

### Environment Variables

Create a `.env` file:

```bash
cp .env.example .env
```

Fill in:

```
AZURE_TENANT_ID=<your-tenant-uuid>
AZURE_CLIENT_ID=<your-app-registration-id>
AZURE_CLIENT_SECRET=<your-client-secret>
# Optional: restrict to specific teams (comma-separated)
# TEAM_IDS=team-id-1,team-id-2
```

---

## Option A: stdio mode (VS Code / Claude Desktop / Claude Code)

### A1. Run from source (development)

```bash
git clone https://github.com/yc-amerpus/teams-mcp.git
cd teams-mcp
npm install
npm run build
```

#### VS Code

Add to `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "teams-mcp": {
      "command": "node",
      "args": ["/path/to/teams-mcp/dist/index.js"],
      "envFile": "/path/to/teams-mcp/.env"
    }
  }
}
```

Then: `Ctrl+Shift+P` > **MCP: Restart Server** > `teams-mcp`

#### Claude Desktop

Add to your Claude Desktop config (`~/Library/Application Support/Claude/claude_desktop_config.json` on Mac, `%APPDATA%\Claude\claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "node",
      "args": ["/path/to/teams-mcp/dist/index.js"],
      "env": {
        "AZURE_TENANT_ID": "<your-tenant-uuid>",
        "AZURE_CLIENT_ID": "<your-app-registration-id>",
        "AZURE_CLIENT_SECRET": "<your-client-secret>"
      }
    }
  }
}
```

#### Claude Code

Add to `~/.claude.json` or use the CLI:

```bash
claude mcp add teams-mcp node /path/to/teams-mcp/dist/index.js \
  -e AZURE_TENANT_ID=<your-tenant-uuid> \
  -e AZURE_CLIENT_ID=<your-app-registration-id> \
  -e AZURE_CLIENT_SECRET=<your-client-secret>
```

### A2. Run via Docker (stdio)

```bash
cd teams-mcp
docker compose -f docker-compose.stdio.yml build
```

Then reference the Docker image from your MCP client:

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "docker",
      "args": ["run", "-i", "--rm", "--env-file", "/path/to/teams-mcp/.env", "teams-mcp:latest"]
    }
  }
}
```

---

## Option B: SSE/HTTP mode (AnythingLLM / remote clients)

This mode uses [supergateway](https://github.com/supercorp-ai/supergateway) to expose the MCP server over HTTP with Server-Sent Events (SSE).

### B1. Docker Compose (recommended)

```bash
cd teams-mcp
cp .env.example .env    # fill in Azure credentials
docker compose up -d
```

The server will be available at `http://localhost:8000/sse`.

#### Connect from AnythingLLM

In AnythingLLM settings, add an MCP server with:
- **Type**: SSE
- **URL**: `http://localhost:8000/sse` (or your host's IP/domain)

#### Connect from a remote client

If deploying on a server, expose port 8000 via a reverse proxy (Traefik, nginx, etc.) with TLS:

```yaml
# Example Traefik dynamic config
http:
  routers:
    teams-mcp:
      rule: "Host(`teams-mcp.example.com`)"
      entryPoints:
        - websecure
      tls:
        certResolver: le-dns
      service: teams-mcp

  services:
    teams-mcp:
      loadBalancer:
        servers:
          - url: "http://teams-mcp:8000"
```

### B2. Run without Docker

```bash
npm install
npm run build
npm install -g supergateway

# Load env vars and start
export $(cat .env | xargs)
supergateway --stdio "node dist/index.js" --port 8000 --host 0.0.0.0
```

---

## Verify

### Quick smoke test (stdio)

```bash
export $(cat .env | xargs)
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}
{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}' | node dist/index.js 2>/dev/null
```

Should return JSON listing 15 available tools.

### Quick smoke test (SSE)

```bash
curl http://localhost:8000/sse
```

Should return an SSE event stream.

### Test auth

Call the `check_auth` tool — it should return your tenant and client IDs.

### Test teams

Call `list_teams` — it should return your configured teams.

---

## Available Tools

| Tool | Description |
|---|---|
| `check_auth` | Verify authentication status |
| `get_app_info` | Show app permissions, capabilities, limitations |
| `search_users` | Search users by name or email |
| `get_user` | Get user profile by ID or email |
| `list_teams` | List accessible teams |
| `list_channels` | List channels in a team |
| `get_channel_messages` | Read channel messages with metadata |
| `send_channel_message` | Send a message (text or markdown) |
| `reply_to_channel_message` | Reply to a message |
| `get_channel_message_replies` | Get replies to a message |
| `list_team_members` | List team members with roles |
| `search_users_for_mentions` | Search users for @mention formatting |
| `download_message_hosted_content` | Download images/files from messages |
| `update_channel_message` | Edit a channel message |
| `delete_channel_message` | Soft-delete a channel message |

## Troubleshooting

### "Missing required environment variables"

Ensure `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` are set. Check your `.env` file has no extra spaces or quotes.

### "Insufficient privileges" on check_auth

The app registration is authenticated but missing Graph API permissions. Go to Azure Portal > App Registration > API Permissions and grant the required application permissions with admin consent.

### "No teams found matching TEAM_IDS filter"

Either the `TEAM_IDS` value doesn't match any team, or the app lacks `Team.ReadBasic.All` permission. Try unsetting `TEAM_IDS` to list all teams in the tenant.

### jsdom ESM error on Node 18

If you see `require() of ES Module encoding-lite.js not supported`, ensure you're using jsdom v24 (already pinned in package.json). Node 20+ does not have this issue.

### Docker: "Permission denied"

Ensure the `.env` file is readable and the Docker daemon is running:

```bash
ls -la .env
docker info
```
