# CLAUDE.md

## Project Overview

Teams-MCP is a Model Context Protocol (MCP) server providing Microsoft Teams integration via the Microsoft Graph API. It uses **app-only client credentials authentication** (no interactive login, runs headlessly). Forked from [floriscornel/teams-mcp](https://github.com/floriscornel/teams-mcp) and converted from delegated OAuth to client credentials grant.

## Tech Stack

- **Language:** TypeScript (strict mode), ES Modules
- **Runtime:** Node.js 18+
- **MCP SDK:** `@modelcontextprotocol/sdk`
- **Auth:** `@azure/msal-node` (ConfidentialClientApplication, client credentials flow)
- **API:** `@microsoft/microsoft-graph-client`
- **Markdown:** `marked` + `dompurify` + `jsdom` (v24 for Node 18 compat)
- **Validation:** `zod`
- **Linter/Formatter:** Biome
- **Testing:** Vitest (80% coverage threshold) + sinon + msw + nock
- **CI:** GitHub Actions (Node 20/22/24 matrix)

## Commands

```bash
npm install          # Install dependencies
npm run build        # Compile TypeScript to dist/
npm run dev          # Hot reload with tsx watch
npm test             # Run tests
npm test:watch       # Watch mode
npm test:coverage    # Coverage report
npm run lint         # Biome lint check
npm run lint:fix     # Auto-fix lint issues
npm run format       # Format with Biome
npm run ci           # Biome CI check (used in CI/CD)
```

## Project Structure

```
src/
├── index.ts              # Entry point: env validation, MCP server init, tool registration
├── services/
│   └── graph.ts          # GraphService singleton: MSAL auth, token management, Graph API
├── tools/
│   ├── auth.ts           # check_auth, get_app_info
│   ├── teams.ts          # Teams/channels CRUD (list, read, send, reply, edit, delete)
│   └── users.ts          # User search, profiles, mention formatting
├── types/
│   └── graph.ts          # TypeScript interfaces for Graph API responses
├── utils/
│   ├── markdown.ts       # Markdown→HTML with DOMPurify sanitization
│   ├── users.ts          # User search and @mention utilities
│   └── attachments.ts    # Image/file upload and validation
└── test-utils/           # Vitest setup and test helpers
```

## Architecture

- **GraphService** is a singleton (`GraphService.getInstance()`) managing auth and API calls
- **Auth priority:** `AUTH_TOKEN` env var (direct JWT) > MSAL client credentials
- **Tool modules** register via `registerAuthTools()`, `registerTeamsTools()`, `registerUsersTools()`
- **Deployment modes:** stdio (native MCP) or SSE/HTTP (via supergateway bridge)
- **Docker:** `Dockerfile` for stdio, `Dockerfile.sse` for SSE mode

## Environment Variables

Required: `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`
Optional: `TEAM_IDS` (comma-separated filter), `AUTH_TOKEN` (direct JWT for testing)

## Code Style

- Biome enforced: 2-space indent, double quotes, semicolons, ES5 trailing commas, 100 char line width
- Test files have relaxed arrow function rules

## App-Only Limitations

Cannot support: 1:1/group chat (protected API), message search (delegated-only), `/me` endpoint (no user context).
