import { createHash, randomBytes } from 'node:crypto';
import { InvalidGrantError, InvalidRequestError, InvalidTokenError } from '@modelcontextprotocol/sdk/server/auth/errors.js';
import { config } from './standard-config.mjs';
import { createConnection, startMicrosoftLogin } from './standard-microsoft.mjs';

const opaque = () => randomBytes(32).toString('base64url');
const digest = value => createHash('sha256').update(value).digest('hex');
const escape = value => String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const page = body => `<!doctype html><html lang="zh"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Microsoft Copilot MCP</title><style>body{font:17px/1.6 system-ui;max-width:680px;margin:60px auto;padding:24px;color:#172b3a}button,a{font:inherit}button{padding:12px 18px;background:#245d70;color:white;border:0;border-radius:8px}code{font-size:24px}small{color:#556}</style><h1>Microsoft Copilot MCP</h1>${body}</html>`;

export function createOAuthProvider(origin, { login = startMicrosoftLogin, connectionOptions = {} } = {}) {
  // Keep the browser OAuth origin explicit. It must match the public origin so
  // the OAuth flow cookie remains on the same host after Microsoft sign-in.
  const browserOrigin = process.env.MCP_BROWSER_ORIGIN ?? origin;
  const resource = `${origin}/mcp`;
  const clients = new Map();
  if (config.mcpClientId) clients.set(config.mcpClientId, { client_id: config.mcpClientId, redirect_uris: config.redirects, token_endpoint_auth_method: 'none', grant_types: ['authorization_code', 'refresh_token'], response_types: ['code'], client_name: 'Configured MCP client' });
  const flows = new Map(), codes = new Map(), tokens = new Map(), refreshes = new Map(), connections = new Map();
  const expires = (ms = 8 * 60 * 60 * 1000) => Math.min(Date.now() + ms, Date.parse(config.expiresAt));
  const validateResource = value => { if (value && value.toString() !== resource) throw new InvalidRequestError('Unexpected resource'); };
  const issue = (clientId, connectionId) => {
    const access_token = opaque(); const refresh_token = opaque();
    tokens.set(digest(access_token), { clientId, connectionId, expires: expires(3600000) });
    refreshes.set(digest(refresh_token), { clientId, connectionId, expires: expires() });
    return { access_token, refresh_token, token_type: 'Bearer', expires_in: Math.max(0, Math.floor((expires(3600000) - Date.now()) / 1000)), scope: 'copilot.chat' };
  };
  const validCode = (client, code) => { const item = codes.get(digest(code)); if (!item || item.expires < Date.now() || item.clientId !== client.client_id) throw new InvalidGrantError('Invalid or expired authorization code'); return item; };
  function sweep() {
    for (const map of [flows, codes, tokens, refreshes]) for (const [key, value] of map) if (value.expires < Date.now()) map.delete(key);
    for (const [key, value] of connections) if (value.expires < Date.now()) { void value.connection.disconnect(); connections.delete(key); }
  }
  const provider = {
    clientsStore: {
      async getClient(id) { return clients.get(id); },
      async registerClient(metadata) {
        if (clients.size >= 20 || !metadata.redirect_uris?.length || metadata.redirect_uris.some(uri => !config.redirects.includes(uri))) throw new InvalidRequestError('Only configured MCP callback URLs are allowed');
        // Return the configured identity instead of creating an in-memory ID
        // that changes after every process restart.
        const configured = clients.get(config.mcpClientId);
        if (configured) return configured;
        const client = { ...metadata, client_id: config.mcpClientId, token_endpoint_auth_method: 'none', client_secret: undefined, grant_types: ['authorization_code', 'refresh_token'], response_types: ['code'] };
        clients.set(client.client_id, client); return client;
      },
    },
    async authorize(client, params, res) {
      sweep(); validateResource(params.resource);
      if (!config.redirects.includes(params.redirectUri) || !/^[A-Za-z0-9_-]{43}$/.test(params.codeChallenge)) throw new InvalidRequestError('Invalid redirect or PKCE challenge');
      const id = opaque(); const cookie = opaque();
      flows.set(id, { ...params, clientId: client.client_id, cookie: digest(cookie), csrf: opaque(), expires: expires(600000), stage: 'ready' });
      res.cookie('mcp_flow', cookie, { httpOnly: true, secure: browserOrigin.startsWith('https:'), sameSite: 'lax', maxAge: 600000, path: '/connect' });
      res.redirect(`${browserOrigin}/connect?flow=${encodeURIComponent(id)}`);
    },
    async challengeForAuthorizationCode(client, code) { return validCode(client, code).codeChallenge; },
    async exchangeAuthorizationCode(client, code, verifier, redirectUri, target) {
      validateResource(target); const item = validCode(client, code);
      if (redirectUri !== item.redirectUri) throw new InvalidGrantError('Redirect mismatch');
      codes.delete(digest(code)); return issue(client.client_id, item.connectionId);
    },
    async exchangeRefreshToken(client, token, scopes, target) {
      validateResource(target);
      const item = refreshes.get(digest(token));
      if (!item || item.clientId !== client.client_id || item.expires < Date.now() || !connections.has(item.connectionId)) throw new InvalidGrantError('Invalid refresh token');
      refreshes.delete(digest(token)); return issue(client.client_id, item.connectionId);
    },
    async verifyAccessToken(token) {
      sweep(); const item = tokens.get(digest(token));
      if (!item || !connections.has(item.connectionId)) throw new InvalidTokenError('Invalid or expired MCP token');
      return { token, clientId: item.clientId, scopes: ['copilot.chat'], expiresAt: Math.floor(item.expires / 1000), resource: new URL(resource), extra: { connectionId: item.connectionId } };
    },
    async revokeToken(client, { token }) {
      const item = tokens.get(digest(token)) ?? refreshes.get(digest(token));
      if (!item || item.clientId !== client.client_id) return;
      for (const map of [tokens, refreshes]) for (const [key, value] of map) if (value.connectionId === item.connectionId) map.delete(key);
      await connections.get(item.connectionId)?.connection.disconnect(); connections.delete(item.connectionId);
    },
    getConnection(id) { const entry = connections.get(id); if (!entry || entry.expires < Date.now()) throw new InvalidTokenError('MCP connection missing'); return entry.connection; },
    installUI(app, formParser) {
      const flowFor = req => {
        const id = req.query.flow;
        const flow = typeof id === 'string' ? flows.get(id) : undefined;
        const cookie = req.headers.cookie?.split(';').map(value => value.trim()).find(value => value.startsWith('mcp_flow='))?.slice(9);
        if (!flow || flow.expires < Date.now() || !cookie || digest(cookie) !== flow.cookie) throw new Error('INVALID_FLOW');
        return { id, flow };
      };
      const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
      const startLogin = (id, flow) => {
        if (flow.loginPromise || flow.stage !== 'ready') return flow.loginReady;
        let resolveReady;
        flow.loginReady = new Promise(resolve => { resolveReady = resolve; });
        flow.stage = 'pending';
        flow.loginPromise = Promise.resolve(login(device => {
          flow.device = device;
          resolveReady(device);
        }, { successRedirectUrl: `${browserOrigin}/connect?flow=${encodeURIComponent(id)}` }))
          .then(identity => {
            if (flow.expires >= Date.now()) {
              flow.identity = identity;
              flow.stage = 'authenticated';
            }
            return identity;
          })
          .catch(error => {
            flow.loginError = error instanceof Error ? error.message : 'MICROSOFT_LOGIN_FAILED';
            flow.stage = 'failed';
            throw error;
          });
        void flow.loginPromise.catch(() => {});
        return flow.loginReady;
      };
      const finish = (id, flow, res) => {
        if (flow.stage !== 'authenticated' || !flow.identity) throw new Error('INVALID_STAGE');
        flow.stage = 'completed';
        const connection = createConnection(flow.identity, connectionOptions);
        connections.set(connection.id, { connection, expires: expires() });
        const code = opaque();
        codes.set(digest(code), { clientId: flow.clientId, connectionId: connection.id, codeChallenge: flow.codeChallenge, redirectUri: flow.redirectUri, expires: expires(120000) });
        const target = new URL(flow.redirectUri);
        target.searchParams.set('code', code);
        if (flow.state) target.searchParams.set('state', flow.state);
        flows.delete(id);
        res.clearCookie('mcp_flow', { path: '/connect' });
        res.redirect(target.toString());
      };
      const renderFallback = (id, flow, res) => {
        const retry = `${browserOrigin}/connect?flow=${encodeURIComponent(id)}`;
        let body = '<p>正在连接 Microsoft 账号，请稍候。</p>';
        if (flow.device?.verificationUri) body = `<p>请打开 Microsoft 登录页面并输入设备码：</p><p><code>${escape(flow.device.userCode)}</code></p><p><a href="${escape(flow.device.verificationUri)}" target="_blank" rel="noreferrer">打开 Microsoft 登录</a></p>`;
        if (flow.loginError) body = '<p>Microsoft 登录没有完成，请返回 MCP 客户端重新连接。</p>';
        res.type('html').send(page(`${body}<meta http-equiv="refresh" content="2;url=${escape(retry)}">`));
      };
      app.get('/connect', async (req, res) => {
        try {
          const { id, flow } = flowFor(req);
          if (flow.stage === 'ready') {
            const ready = startLogin(id, flow);
            await Promise.race([ready, wait(3000)]);
          }
          if (flow.stage === 'authenticated') return finish(id, flow, res);
          if (flow.stage === 'pending' && flow.device?.authorizationUrl) return res.redirect(flow.device.authorizationUrl);
          renderFallback(id, flow, res);
        } catch { if (!res.headersSent) res.status(400).send('Invalid or expired connection flow.'); }
      });
      app.post('/connect', formParser, (req, res) => {
        try {
          const { id, flow } = flowFor(req);
          if (![origin, browserOrigin].includes(req.headers.origin) || req.body.csrf !== flow.csrf) throw new Error('INVALID_CSRF');
          if (req.body.action === 'start' && flow.stage === 'ready') {
            startLogin(id, flow);
            res.redirect(303, `${browserOrigin}/connect?flow=${encodeURIComponent(id)}`); return;
          } else if (req.body.action === 'finish' && flow.stage === 'authenticated') {
            finish(id, flow, res); return;
          } else throw new Error('INVALID_STAGE');
          res.redirect(303, `${browserOrigin}/connect?flow=${id}`);
        } catch { res.status(400).send('Invalid connection action.'); }
      });
    },
    async close() { for (const { connection } of connections.values()) await connection.disconnect(); connections.clear(); tokens.clear(); refreshes.clear(); flows.clear(); codes.clear(); },
  };
  return provider;
}
