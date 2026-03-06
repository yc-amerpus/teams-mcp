# Teams MCP

A Model Context Protocol (MCP) server for Microsoft Teams integration via the Microsoft Graph API. Uses **app-only client credentials** authentication — no interactive login, no browser, runs headlessly as a service.

Forked from [floriscornel/teams-mcp](https://github.com/floriscornel/teams-mcp) and converted from delegated OAuth (device code flow) to client credentials grant.

## Quick Start

```bash
git clone https://github.com/yc-amerpus/teams-mcp.git
cd teams-mcp
npm install && npm run build
cp .env.example .env   # fill in Azure credentials
```

**VS Code** — add to `.vscode/mcp.json`:

```json
{
  "servers": {
    "teams-mcp": {
      "command": "node",
      "args": ["dist/index.js"],
      "envFile": "${workspaceFolder}/.env"
    }
  }
}
```

**SSE/HTTP** (for AnythingLLM, remote clients):

```bash
docker compose up -d
# Server available at http://localhost:8000/sse
```

See [INSTALL.md](INSTALL.md) for full deployment instructions including Docker, Claude Desktop, Claude Code, and Traefik reverse proxy setup.

## Features

### Authentication
- Client credentials grant (app-only) via MSAL `ConfidentialClientApplication`
- Reads `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` from environment
- No interactive auth, no token cache files, no browser
- MSAL handles token refresh internally

### Teams & Channels
- List teams (all tenant teams, or filtered by `TEAM_IDS` env var)
- List channels within teams
- Read channel messages with full metadata (author, timestamps, importance)
- Send messages and replies (text or markdown with HTML sanitization)
- Edit and soft-delete channel messages
- Download hosted content (images, files) from messages

### Users & Directory
- Search users by name or email
- Get detailed user profiles
- Search users for @mention formatting

## Available Tools

| Tool | Description |
|---|---|
| `check_auth` | Verify authentication status |
| `get_app_info` | Show permissions, capabilities, and limitations |
| `list_teams` | List accessible teams |
| `list_channels` | List channels in a team |
| `get_channel_messages` | Read messages with author and timestamps |
| `send_channel_message` | Send a message (text or markdown) |
| `reply_to_channel_message` | Reply to a channel message |
| `get_channel_message_replies` | Get replies to a message |
| `update_channel_message` | Edit a sent message |
| `delete_channel_message` | Soft-delete a message |
| `list_team_members` | List members with roles |
| `search_users` | Search directory by name or email |
| `get_user` | Get user profile by ID or email |
| `search_users_for_mentions` | Find users for @mentions |
| `download_message_hosted_content` | Download images/files from messages |

## Azure App Registration

Required **application permissions** (not delegated), with admin consent:

| Permission | Purpose |
|---|---|
| `Team.ReadBasic.All` | List teams |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Send/reply to messages |
| `TeamMember.Read.All` | List team members |
| `User.Read.All` | Search/get user profiles |

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `AZURE_TENANT_ID` | Yes | Azure AD tenant UUID |
| `AZURE_CLIENT_ID` | Yes | App registration client ID |
| `AZURE_CLIENT_SECRET` | Yes | App registration client secret |
| `TEAM_IDS` | No | Comma-separated team IDs to filter |
| `AUTH_TOKEN` | No | Direct token injection (bypasses MSAL) |

## Deployment Modes

### stdio (VS Code, Claude Desktop, Claude Code)

The server communicates over stdin/stdout using the MCP protocol. Configure your MCP client to run `node dist/index.js` with the required environment variables.

### SSE/HTTP (AnythingLLM, remote clients)

Uses [supergateway](https://github.com/supercorp-ai/supergateway) to bridge stdio to HTTP/SSE:

```bash
docker compose up -d
# or without Docker:
supergateway --stdio "node dist/index.js" --port 8000 --host 0.0.0.0
```

## Markdown Support

Messages sent via `send_channel_message`, `reply_to_channel_message`, and `update_channel_message` support a `format` parameter:

- `text` (default) — plain text
- `markdown` — converted to sanitized HTML (bold, italic, links, lists, code blocks, headings, tables)

## What's Not Included

These features from the original floriscornel/teams-mcp require delegated permissions or protected APIs and are not available in app-only mode:

| Feature | Reason |
|---|---|
| Chat (1:1 and group) | Protected API — requires Microsoft approval |
| Message search | Microsoft Search API is delegated-only |
| Current user context | No `/me` endpoint in app-only flow |

## Development

```bash
npm install
npm run dev          # hot reload with tsx
npm run build        # compile TypeScript
npm test             # run tests
npm run lint         # biome check
```

## License

MIT License — see LICENSE file for details.
