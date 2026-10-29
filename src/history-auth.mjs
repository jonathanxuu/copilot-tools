import { readFileSync } from 'node:fs';
import { createPrivateKey } from 'node:crypto';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { config, graphBase } from './standard-config.mjs';
import { createInteractionHistoryReader } from './interaction-history.mjs';

export function createHistoryReaderFactory({ env = process.env } = {}) {
  if (env.M365_MCP_HISTORY_ENABLED !== 'true') return undefined;
  const keyPath = env.M365_MCP_HISTORY_PRIVATE_KEY_PATH;
  const thumbprintSha256 = (env.M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256 ?? '').replace(/[:\s]/g, '');
  const defaultDays = Number(env.M365_MCP_HISTORY_DEFAULT_DAYS ?? 30);
  const maxRequests = Number(env.M365_MCP_HISTORY_MAX_REQUESTS ?? 10);
  const allowedSessionIds = env.M365_MCP_HISTORY_ALLOWED_SESSION_IDS ? env.M365_MCP_HISTORY_ALLOWED_SESSION_IDS.split(',').map(value => value.trim()).filter(Boolean) : undefined;
  if (!keyPath || !/^[a-f0-9]{64}$/i.test(thumbprintSha256) || !Number.isInteger(defaultDays) || defaultDays < 1 || defaultDays > 365 || !Number.isInteger(maxRequests) || maxRequests < 1 || maxRequests > 32) throw new Error('CONFIG_HISTORY_REQUIRED');
  const privateKey = readFileSync(keyPath, 'utf8');
  if (!/-----BEGIN (?:RSA )?PRIVATE KEY-----/.test(privateKey)) throw new Error('CONFIG_HISTORY_PRIVATE_KEY_INVALID');
  createPrivateKey(privateKey);
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const clients = new Map();
  const clientFor = tenantId => {
    if (!guid.test(tenantId ?? '')) throw new Error('HISTORY_TENANT_REQUIRED');
    if (!clients.has(tenantId)) clients.set(tenantId, new ConfidentialClientApplication({ auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${tenantId}`, clientCertificate: { thumbprintSha256, privateKey } }, system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } } }));
    return clients.get(tenantId);
  };
  async function getToken(tenantId) {
    const client = clientFor(tenantId);
    const result = await client.acquireTokenByClientCredential({ scopes: [`${graphBase}/.default`] });
    if (!result?.accessToken) throw new Error('HISTORY_TOKEN_UNAVAILABLE');
    return result.accessToken;
  }
  const budget = { used: 0, max: maxRequests };
  return identity => {
    const to = Date.now();
    const from = to - defaultDays * 86400000;
    return createInteractionHistoryReader({ tenantId: identity.tenantId, userId: identity.userId, boundSubject: identity.subject, from: new Date(from).toISOString(), to: new Date(to).toISOString(), maxLookbackMs: 5 * 366 * 86400000, deadline: Date.parse(config.expiresAt), maxRequests, budget, allowedSessionIds, getToken: () => getToken(identity.tenantId), onAudit: event => console.log(JSON.stringify({ time: new Date().toISOString(), tenantId: identity.tenantId, userId: identity.userId, ...event })) });
  };
}
