# Microsoft Copilot MCP Server

This is a standalone MCP server that exposes Microsoft Copilot capabilities to any MCP-compatible client. The client can be Dust, Claude, Cursor, an internal agent, or another MCP host; the server does not call or depend on any client's API.

The runtime now uses the MCP SDK OAuth provider for MCP-client authorization and Microsoft MSAL for Microsoft sign-in. These are separate layers: the first authenticates the MCP client, the second obtains the user's delegated Microsoft token.

The initial version is intentionally read-only:

- `get_microsoft_capabilities`
- `microsoft_copilot_chat`
- `microsoft_list_history`
- `microsoft_read_copilot_history`

It does not write to an MCP client and does not share a Microsoft token between users.

## Local setup

```bash
cp .env.example .env
npm install
npm run check
npm run dev
```

Health check:

```bash
curl http://localhost:8787/healthz
```

MCP endpoint:

```text
http://localhost:8787/mcp
```

## MCP client setup

Deploy the service to `https://copilot-mcp.zkid.xyz` (or another HTTPS origin), then add the `/mcp` URL to the MCP client's remote-server configuration. For a multi-tenant deployment, use:

```env
MCP_PUBLIC_ORIGIN=https://copilot-mcp.zkid.xyz
MCP_BROWSER_ORIGIN=https://copilot-mcp.zkid.xyz
MCP_BIND_HOST=0.0.0.0
M365_MCP_TENANT_ID=organizations
M365_MCP_LOGIN_MODE=browser
```

`MCP_PUBLIC_ORIGIN` is the origin advertised to remote MCP clients. `MCP_BROWSER_ORIGIN` is the origin used for the browser OAuth flow; keeping them identical prevents OAuth cookies from being split across hostnames. `organizations` allows work/school accounts from multiple Entra tenants; customer administrators still need to consent to the app. Use `browser` for a hosted HTTPS service: Microsoft redirects the user to `M365_MCP_BROWSER_REDIRECT_URI` on this server. `device-code` is retained only for legacy environments because customer tenants may block device-code flow with Security Defaults.

The server exposes standard MCP OAuth endpoints:

```text
Authorization server metadata: https://your-mcp.example.com/.well-known/oauth-authorization-server
Dynamic client registration:   https://your-mcp.example.com/oauth/register
Authorization endpoint:        https://your-mcp.example.com/oauth/authorize
Token endpoint:                https://your-mcp.example.com/oauth/token
MCP endpoint:                  https://your-mcp.example.com/mcp
```

An MCP client registers its callback through the MCP SDK OAuth router, uses PKCE S256, and then opens the connection page. The connection page starts a Microsoft MSAL browser authorization-code login. Register the exact value of `M365_MCP_BROWSER_REDIRECT_URI` as a Web redirect URI in the Entra app registration. Add the MCP client's callback URL to `M365_MCP_REDIRECT_URIS`. Do not configure a shared Microsoft token.

The hosted browser flow uses a confidential Microsoft client. Configure either `M365_MCP_ENTRA_CLIENT_SECRET`, or `M365_MCP_ENTRA_PRIVATE_KEY_PATH` together with `M365_MCP_ENTRA_CERT_THUMBPRINT_SHA256`. If Interaction Export is enabled, the browser flow can reuse `M365_MCP_HISTORY_PRIVATE_KEY_PATH` and `M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256` when the same certificate is registered on the app.

## Important configuration notes

Microsoft Copilot history and Pages are separate Microsoft capabilities. History uses the Interaction Export API with app-only certificate permission. Pages are currently disabled in the MCP tool list because their SharePoint Embedded container access requires additional app-only guest application configuration.

The current Chat implementation uses the Microsoft 365 Copilot Chat API preview shape:

```text
POST /beta/copilot/conversations
POST /beta/copilot/conversations/{conversationId}/chat
```

The Microsoft adapter is isolated in `src/standard-microsoft.mjs` so endpoint or response-shape changes do not affect the client-facing MCP tools.

## Authentication and storage

The current implementation uses the MCP SDK OAuth provider and keeps MCP connections in memory. Microsoft MSAL cache data is encrypted at rest using `M365_MCP_TOKEN_CACHE_KEY` and stored at `M365_MCP_TOKEN_CACHE_PATH`; the MCP bearer connection still expires and requires re-authorization after a restart. Copilot conversation-handle mappings and successful real-time Chat API results are persisted at `M365_MCP_CONVERSATION_STATE_PATH`, so a caller that reuses the returned `conversationHandle` can continue the same Copilot conversation after a server restart and can inspect local activity before Microsoft Interaction Export includes it.

Copilot history is disabled unless `M365_MCP_HISTORY_ENABLED=true`. When enabled, the server uses a certificate-backed app-only Microsoft Graph token with `AiEnterpriseInteraction.Read.All`, then locally binds results to the signed-in user's tenant and object ID. History queries default to the most recent 30 days; callers may provide custom dates within the server-enforced five-year maximum lookback. `M365_MCP_HISTORY_DEFAULT_DAYS` changes the default window.

History can be further restricted with `M365_MCP_HISTORY_ALLOWED_SESSION_IDS`. The list tool also accepts `sessionIds` to narrow the returned sessions; a session must be listed before the read tool will return its messages. History queries and tool outcomes are written to audit logs without access tokens or message bodies.

Copilot Pages adapter code is retained in `src/standard-microsoft.mjs`, but its MCP tools are intentionally disabled for now. Re-enable them only after SharePoint Embedded app-only guest access is granted to the MCP Entra application, an app-only token flow is implemented, and the user-container binding is tested. See `TODO.md`. The imported content remains intended to be a private in-memory context of the current MCP connection; it is not written to Dust or another external store.

## Security rules

- Never log access tokens or authorization codes.
- Keep tenant ID, Microsoft user ID, and MCP client ID in audit records.
- Keep Page preview read-only until the content visibility model is agreed.
- Add a client-specific save tool only after testing explicit confirmation and private-vs-shared visibility.

## Personal Microsoft accounts

Personal accounts may be able to complete basic Microsoft sign-in, but Microsoft 365 Copilot Chat, enterprise interaction history, and Copilot Pages routes used by this project require a supported work/school Microsoft 365 tenant and licensing. Use `MICROSOFT_ACCOUNT_MODE=enterprise` only after that tenant is configured.
