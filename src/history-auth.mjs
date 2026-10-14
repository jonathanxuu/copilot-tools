import { readFileSync } from 'node:fs';
import { createPrivateKey } from 'node:crypto';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { config, graphBase } from './standard-config.mjs';
import { createInteractionHistoryReader } from './interaction-history.mjs';

export function createHistoryReaderFactory({ env = process.env } = {}) {
  if (env.M365_MCP_HISTORY_ENABLED !== 'true') return undefined;
  const keyPath = env.M365_MCP_HISTORY_PRIVATE_KEY_PATH;
  const thumbprintSha256 = (env.M365_MCP_HISTORY_CERT_THUMBPRINT_SHA256 ?? '').replace(/[:\s]/g, '');
  const from = env.M365_MCP_HISTORY_FROM;
  const to = env.M365_MCP_HISTORY_TO;
  const maxRequests = Number(env.M365_MCP_HISTORY_MAX_REQUESTS ?? 10);
  const allowedSessionIds = env.M365_MCP_HISTORY_ALLOWED_SESSION_IDS ? env.M365_MCP_HISTORY_ALLOWED_SESSION_IDS.split(',').map(value => value.trim()).filter(Boolean) : undefined;
  if (!keyPath || !/^[a-f0-9]{64}$/i.test(thumbprintSha256) || !from || !to || !Number.isInteger(maxRequests) || maxRequests < 1 || maxRequests > 32) throw new Error('CONFIG_HISTORY_REQUIRED');
  const start = Date.parse(from), end = Date.parse(to);
  if (!Number.isFinite(start) || !Number.isFinite(end) || start >= end || end - start > 5 * 366 * 86400000) throw new Error('CONFIG_HISTORY_WINDOW_INVALID');
  const privateKey = readFileSync(keyPath, 'utf8');
  if (!/-----BEGIN (?:RSA )?PRIVATE KEY-----/.test(privateKey)) throw new Error('CONFIG_HISTORY_PRIVATE_KEY_INVALID');
  createPrivateKey(privateKey);
  const client = new ConfidentialClientApplication({ auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, clientCertificate: { thumbprintSha256, privateKey } }, system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } } });
  async function getToken() {
    const result = await client.acquireTokenByClientCredential({ scopes: [`${graphBase}/.default`] });
    if (!result?.accessToken) throw new Error('HISTORY_TOKEN_UNAVAILABLE');
    return result.accessToken;
  }
  const budget = { used: 0, max: maxRequests };
  return identity => createInteractionHistoryReader({ tenantId: config.tenantId, userId: identity.userId, boundSubject: identity.subject, from, to, deadline: Date.parse(config.expiresAt), maxRequests, budget, allowedSessionIds, getToken, onAudit: event => console.log(JSON.stringify({ time: new Date().toISOString(), userId: identity.userId, ...event })) });
}
