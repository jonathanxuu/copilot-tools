const list = value => (value ?? '').split(',').map(part => part.trim()).filter(Boolean);

export const config = Object.freeze({
  tenantId: process.env.M365_MCP_TENANT_ID ?? process.env.MICROSOFT_TENANT_ID ?? '',
  clientId: process.env.M365_MCP_ENTRA_CLIENT_ID ?? process.env.MICROSOFT_CLIENT_ID ?? '',
  allowedUsername: (process.env.M365_MCP_ALLOWED_USERNAME ?? '').toLowerCase(),
  // Keep one stable MCP client identity across restarts.
  mcpClientId: process.env.M365_MCP_OAUTH_CLIENT_ID || 'microsoft-copilot-mcp',
  conversationStatePath: process.env.M365_MCP_CONVERSATION_STATE_PATH ?? './data/conversations.json',
  redirects: list(process.env.M365_MCP_REDIRECT_URIS),
  timeZone: process.env.M365_MCP_TIME_ZONE ?? 'Asia/Shanghai',
  expiresAt: process.env.M365_MCP_EXPIRES_AT ?? '2099-01-01T00:00:00Z',
  graphScopes: list(process.env.M365_MCP_GRAPH_SCOPES ?? 'User.Read,Sites.Read.All,Mail.Read,People.Read.All,OnlineMeetingTranscript.Read.All,Chat.Read,ChannelMessage.Read.All,ExternalItem.Read.All,openid,profile,offline_access'),
  loginMode: process.env.M365_MCP_LOGIN_MODE ?? 'loopback',
  maxGraphRequests: Number(process.env.M365_MCP_MAX_GRAPH_REQUESTS ?? 12),
  maxPageBytes: Number(process.env.M365_MCP_MAX_PAGE_BYTES ?? 2 * 1024 * 1024),
});

export const graphBase = (process.env.MICROSOFT_GRAPH_BASE_URL ?? 'https://graph.microsoft.com').replace(/\/$/, '');

export function isAllowedRedirect(uri) {
  try {
    const url = new URL(uri);
    return !url.username && !url.password && !url.hash && (url.protocol === 'https:' || (url.protocol === 'http:' && ['127.0.0.1', 'localhost', '[::1]'].includes(url.hostname)));
  } catch { return false; }
}

export function assertRunnableConfig() {
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!config.tenantId || !config.clientId) throw new Error('CONFIG_MICROSOFT_CLIENT_REQUIRED');
  if (process.env.M365_MCP_HISTORY_ENABLED === 'true' && (!guid.test(config.tenantId) || !guid.test(config.clientId))) throw new Error('CONFIG_HISTORY_TENANT_CLIENT_GUID_REQUIRED');
  if (!config.redirects.length || config.redirects.some(uri => !isAllowedRedirect(uri))) throw new Error('CONFIG_MCP_REDIRECT_REQUIRED');
  if (!Number.isInteger(config.maxGraphRequests) || config.maxGraphRequests < 1 || config.maxGraphRequests > 32) throw new Error('CONFIG_REQUEST_CAP_INVALID');
  if (!Number.isInteger(config.maxPageBytes) || config.maxPageBytes < 1024 || config.maxPageBytes > 16 * 1024 * 1024) throw new Error('CONFIG_PAGE_SIZE_INVALID');
  try { new Intl.DateTimeFormat('en', { timeZone: config.timeZone }); } catch { throw new Error('CONFIG_TIME_ZONE_REQUIRED'); }
  if (!Number.isFinite(Date.parse(config.expiresAt)) || Date.now() >= Date.parse(config.expiresAt)) throw new Error('CONFIG_EXPIRY_REQUIRED');
}
