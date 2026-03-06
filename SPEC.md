# Feasibility Assessment & Specification: floriscornel/teams-mcp Client Credentials Fork

## 1. Feasibility Assessment

### Verdict: FEASIBLE — with caveats

Converting floriscornel/teams-mcp from delegated OAuth (device code flow) to client credentials (app-only) auth is technically feasible. The codebase is small (~53 commits, TypeScript, clean architecture), and MSAL for Node.js (`@azure/msal-node`) supports `ConfidentialClientApplication` with `acquireTokenByClientCredential()` as a drop-in replacement for the current `PublicClientApplication` + device code flow.

However, the switch from delegated to application permissions changes what the Graph API allows, and some of floriscornel's tools will need modification or removal.

### What works with application permissions

| Tool | Application permission available | Notes |
|---|---|---|
| `list_teams` | ✅ `Team.ReadBasic.All` | Returns all teams in tenant (not "my" teams) |
| `list_channels` | ✅ `Channel.ReadBasic.All` | Works as-is |
| `get_channel_messages` | ✅ `ChannelMessage.Read.All` | Full metadata (author, timestamps) |
| `send_channel_message` | ✅ `ChannelMessage.Send` | Works as-is |
| `reply_to_channel_message` | ✅ `ChannelMessage.Send` | Works as-is |
| `list_team_members` | ✅ `TeamMember.Read.All` | Works as-is |
| `search_users` | ✅ `User.Read.All` | Works as-is |
| `get_user` | ✅ `User.Read.All` | Works as-is |

### What doesn't work (or works differently)

| Tool | Issue | Resolution |
|---|---|---|
| `get_current_user` | No "current user" in app-only context | Remove or replace with app identity info |
| `list_chats` | `/me/chats` and `/chats` not supported for app permissions | Remove, or use `/users/{id}/chats` with `Chat.Read.All` (requires protected API approval from Microsoft — very hard to get) |
| `get_chat_messages` | Requires `Chat.Read.All` (protected API — needs Microsoft approval) | Remove for initial version, or use RSC `ChatMessage.Read.Chat` if app installed in chat |
| `send_chat_message` | Same protected API issue | Remove for initial version |
| `create_chat` | Same protected API issue | Remove for initial version |
| `search_messages` | Microsoft Search API requires delegated permissions only | Remove — no app-only equivalent exists |
| `get_recent_messages` | Depends on search_messages internally | Remove |
| `get_my_mentions` | No "me" context + depends on search | Remove |
| `authenticate` | Device code flow no longer needed | Replace with env var validation on startup |
| `logout` | Token cache no longer needed | Remove |

### Key constraint: Chat and Search are mostly unavailable

The Microsoft Graph Chat APIs (`/chats`, `/me/chats`) are **protected APIs** under application permissions. Getting access requires submitting a formal request to Microsoft with a business justification, and approval can take weeks. The Microsoft Search API (`/search/query`) is delegated-only — there is no application permission equivalent.

This means the fork will be a **Teams channels + directory server** rather than a full Teams+Chat server. This is still a major improvement over InditexTech because you get author names, timestamps, rich metadata, markdown formatting, member lookups, and multi-channel support.

### RSC alternative for chats (future option)

If you install the Teams App into a specific chat (not just a team), you can use RSC permissions like `ChatMessage.Read.Chat` to read that chat's messages without needing the protected `Chat.Read.All`. This is a possible future enhancement but requires per-chat app installation.

---

## 2. Specification

### 2.1 Project overview

Fork `floriscornel/teams-mcp` and modify the auth layer to use MSAL `ConfidentialClientApplication` with client credentials grant. Remove tools that require delegated permissions or protected APIs. Retain all channel, team, member, and user tools with full metadata.

### 2.2 Auth architecture change

**Current (delegated):**
```
User runs `authenticate` → device code flow → browser sign-in → 
tokens cached at ~/.teams-mcp-token-cache.json → 
PublicClientApplication.acquireTokenByDeviceCode() → 
access token with `scp` (delegated scopes)
```

**Target (app-only):**
```
Container starts → reads AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET from env →
ConfidentialClientApplication.acquireTokenByClientCredential({ 
  scopes: ["https://graph.microsoft.com/.default"] 
}) → access token with `roles` (application permissions)
```

No interactive auth. No token cache file. No browser. Token auto-refreshes via MSAL internally.

### 2.3 Environment variables

```
AZURE_TENANT_ID=<tenant-uuid>         # Required
AZURE_CLIENT_ID=<app-registration-id> # Required  
AZURE_CLIENT_SECRET=<client-secret>   # Required
```

### 2.4 Azure App Registration permissions

All **application** permissions (not delegated), with admin consent:

| Permission | Type | Purpose |
|---|---|---|
| `Team.ReadBasic.All` | Application | List teams |
| `Channel.ReadBasic.All` | Application | List channels |
| `ChannelMessage.Read.All` | Application | Read channel messages (with author/timestamps) |
| `ChannelMessage.Send` | Application | Send/reply to channel messages |
| `TeamMember.Read.All` | Application | List team members |
| `User.Read.All` | Application | Search/get user profiles |

**Note:** `ChannelMessage.Read.All` is a tenant-wide application permission. For scoped access, use RSC permissions (`ChannelMessage.Read.Group`) via Teams App installation per-team, similar to the InditexTech approach. The fork should support both modes.

### 2.5 Retained tools (10)

1. `list_teams` — list teams (note: returns all tenant teams, not "my teams")
2. `list_channels` — list channels in a team
3. `get_channel_messages` — read messages with full metadata
4. `send_channel_message` — send message (text or markdown)
5. `reply_to_channel_message` — reply to a channel message
6. `list_team_members` — list members with roles
7. `search_users` — search directory by name/email
8. `get_user` — get user profile by ID or email
9. `get_app_info` — NEW: return the authenticated app's identity (replaces `get_current_user`)
10. `check_auth` — NEW: verify credentials are valid (replaces `authenticate`)

### 2.6 Removed tools (7)

| Tool | Reason |
|---|---|
| `authenticate` | No interactive auth needed |
| `logout` | No token cache to clear |
| `get_current_user` | No user context in app-only flow |
| `list_chats` | Protected API — requires Microsoft approval |
| `get_chat_messages` | Protected API |
| `send_chat_message` | Protected API |
| `create_chat` | Protected API |
| `search_messages` | Delegated-only (Microsoft Search API) |
| `get_recent_messages` | Depends on search_messages |
| `get_my_mentions` | No "me" context + depends on search |

---

## 3. Task List

### Phase 1: Fork and auth refactor

**Task 1.1 — Fork the repository**
- Fork `floriscornel/teams-mcp` to your GitHub account
- Clone locally, verify it builds (`npm install && npm run build`)

**Task 1.2 — Refactor auth module (`src/auth.ts` or equivalent)**
- Replace `PublicClientApplication` with `ConfidentialClientApplication` from `@azure/msal-node`
- Replace `acquireTokenByDeviceCode()` with `acquireTokenByClientCredential()`
- Scope: `["https://graph.microsoft.com/.default"]`
- Authority: `https://login.microsoftonline.com/{AZURE_TENANT_ID}`
- Remove all token cache file logic (no `~/.teams-mcp-token-cache.json`)
- Remove device code callback/polling logic
- Read `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` from `process.env`
- Validate all three are present on startup, exit with clear error if missing

**Task 1.3 — Update Graph client initialisation**
- Ensure the Graph client uses the token from `acquireTokenByClientCredential()`
- The `getAccessToken()` function (or equivalent) should call `acquireTokenByClientCredential()` each time — MSAL handles caching internally
- Remove any user-specific token refresh logic

**Task 1.4 — Remove CLI auth commands**
- Remove the `authenticate`, `check`, and `logout` CLI subcommands from the entry point (`src/index.ts`)
- The server should just start directly when invoked

### Phase 2: Tool modifications

**Task 2.1 — Remove delegated-only tools**
- Remove tool registrations for: `authenticate`, `logout`, `get_current_user`, `list_chats`, `get_chat_messages`, `send_chat_message`, `create_chat`, `search_messages`, `get_recent_messages`, `get_my_mentions`
- Remove corresponding handler code and any supporting utility functions that are now unused

**Task 2.2 — Fix `list_teams` for app context**
- Delegated flow uses `/me/joinedTeams` — this won't work
- Change to `/teams` endpoint (requires `Team.ReadBasic.All` application permission)
- Or use `/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')` as a fallback
- Consider adding an optional `TEAM_ID` env var filter so it doesn't return every team in the tenant

**Task 2.3 — Verify channel tools work unchanged**
- `get_channel_messages`, `send_channel_message`, `reply_to_channel_message`, `list_channels` should all work with app permissions since they use `/teams/{id}/channels/{id}/messages` endpoints
- Test that message responses include `from.user.displayName` and `createdDateTime` fields

**Task 2.4 — Add `get_app_info` tool**
- New tool that returns the app's client ID, tenant ID, and permission set
- Useful for debugging and for the LLM to understand what it can/can't do

**Task 2.5 — Add `check_auth` tool**
- Attempts to acquire a token and returns success/failure
- Calls a simple Graph endpoint (e.g., `/organization`) to verify permissions

### Phase 3: Configuration and deployment

**Task 3.1 — Update `package.json`**
- Change package name (e.g., `@yourscope/teams-mcp-app` or `teams-mcp-app`)
- Update description to clarify this is the app-only (client credentials) variant
- Verify `@azure/msal-node` is in dependencies (it likely already is)

**Task 3.2 — Update Dockerfile**
- Simplify: no auth step needed, just `npm install -g` the package
- Environment variables via `env_file` in compose
- Supergateway bridge remains the same

**Task 3.3 — Update docker-compose.yml**
- Replace the volume mount for token cache (no longer needed)
- Add `env_file: .env` with the three Azure credentials
- Keep Traefik labels pattern

**Task 3.4 — Create `.env.example`**
```
AZURE_TENANT_ID=<your-tenant-uuid>
AZURE_CLIENT_ID=<your-app-registration-id>
AZURE_CLIENT_SECRET=<your-client-secret>
```

**Task 3.5 — Update README**
- Document the app-only auth model
- List required Azure application permissions
- Remove all references to device code flow, browser auth, token cache
- Document which tools are available and which were removed (and why)

### Phase 4: Testing

**Task 4.1 — Unit test auth module**
- Mock `ConfidentialClientApplication` and verify token acquisition
- Test missing env var handling (should fail gracefully)

**Task 4.2 — Integration test against Graph**
- `list_teams` — verify it returns teams
- `list_channels` — verify for a known team
- `get_channel_messages` — verify messages include author and timestamps
- `send_channel_message` — verify a message appears in the channel
- `list_team_members` — verify member list with names
- `search_users` — verify directory lookup

**Task 4.3 — End-to-end test via supergateway**
- Start the container with supergateway
- Send MCP `initialize` + `tools/list` via curl
- Call `get_channel_messages` and verify response structure

---

## 4. Estimated effort

| Phase | Effort |
|---|---|
| Phase 1: Auth refactor | 2-3 hours |
| Phase 2: Tool modifications | 1-2 hours |
| Phase 3: Config & deployment | 1 hour |
| Phase 4: Testing | 1-2 hours |
| **Total** | **5-8 hours** |

The auth refactor is the core change. Everything else flows from it. The codebase is small and well-structured — this is a straightforward modification, not a rewrite.

---

## 5. Risks and mitigations

| Risk | Mitigation |
|---|---|
| `list_teams` returns hundreds of teams tenant-wide | Add optional `TEAM_IDS` env var to filter to specific teams |
| `ChannelMessage.Read.All` is tenant-wide | Use RSC permissions via Teams App installation for scoped access (can be added as a second auth mode later) |
| `send_channel_message` requires bot registration | Verify the app registration has an associated Azure Bot resource, same as InditexTech setup |
| Future chat access needed | Document the RSC-per-chat path as a roadmap item; the protected API approval path as a fallback |
| MSAL token acquisition adds latency | MSAL caches tokens internally; first call ~500ms, subsequent calls ~0ms until expiry |