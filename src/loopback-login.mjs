import http from 'node:http';
import { createHash, randomBytes } from 'node:crypto';

const pendingCallbacks = new Map();
const completedCallbacks = new Map();
const loopbackServers = new Map();

function callbackResponse(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end(body);
}

function createCallbackHandler(port) {
  return (req, res) => {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    const state = url.searchParams.get('state');
    const pending = state ? pendingCallbacks.get(state) : undefined;
    const completed = state ? completedCallbacks.get(state) : undefined;
    if (req.method === 'GET' && url.pathname === '/' && completed) {
      if (completed.expiresAt > Date.now()) {
        res.writeHead(302, { Location: completed.successRedirectUrl });
        res.end();
        return;
      }
      completedCallbacks.delete(state);
    }
    if (req.method !== 'GET' || url.pathname !== '/' || !pending || pending.consumed) {
      console.warn(JSON.stringify({ event: 'microsoft_callback_rejected', method: req.method, path: url.pathname, hasState: Boolean(state), pendingCallbacks: pendingCallbacks.size, stateMatched: Boolean(pending) }));
      callbackResponse(res, 400, 'Invalid login callback.');
      return;
    }
    pending.consumed = true;
    pendingCallbacks.delete(state);
    if (pending.successRedirectUrl) completedCallbacks.set(state, { successRedirectUrl: pending.successRedirectUrl, expiresAt: Date.now() + 120000 });
    if (url.searchParams.has('error')) {
      callbackResponse(res, 400, 'Microsoft sign-in did not complete. Return to your MCP client.');
      pending.reject(new Error('MICROSOFT_INTERACTIVE_DENIED'));
      return;
    }
    const code = url.searchParams.get('code');
    if (!code) {
      callbackResponse(res, 400, 'Missing authorization code.');
      pending.reject(new Error('MISSING_AUTHORIZATION_CODE'));
      return;
    }
    if (pending.successRedirectUrl) {
      res.writeHead(302, { Location: pending.successRedirectUrl });
      res.end();
    } else {
      callbackResponse(res, 200, 'Microsoft login callback received. Return to the MCP connection page.');
    }
    console.log(JSON.stringify({ event: 'microsoft_callback_accepted' }));
    pending.resolve(code);
    void port;
  };
}

async function ensureLoopbackServers(port) {
  if (loopbackServers.has(port)) return loopbackServers.get(port);
  const handler = createCallbackHandler(port);
  const ipv4 = http.createServer(handler);
  const ipv6 = http.createServer(handler);
  const ready = (async () => {
    await new Promise((resolve, reject) => { ipv4.once('error', reject); ipv4.listen(port, '127.0.0.1', resolve); });
    await new Promise((resolve, reject) => {
      const onError = error => {
        if (['EAFNOSUPPORT', 'EADDRNOTAVAIL', 'ENODEV'].includes(error.code)) resolve();
        else reject(error);
      };
      ipv6.once('error', onError);
      ipv6.listen(port, '::1', resolve);
    });
    return { ipv4, ipv6, port };
  })();
  loopbackServers.set(port, ready);
  try { return await ready; } catch (error) { loopbackServers.delete(port); ipv4.close(); ipv6.close(); throw error; }
}

export async function loopbackLogin(pca, scopes, onReady, { port = 8788, timeoutMs = 300000, successRedirectUrl } = {}) {
  const state = randomBytes(32).toString('base64url');
  const verifier = randomBytes(32).toString('base64url');
  const challenge = createHash('sha256').update(verifier).digest('base64url');
  let resolveCode;
  let rejectCode;
  let consumed = false;
  const result = new Promise((resolve, reject) => { resolveCode = resolve; rejectCode = reject; });
  void result.catch(() => {});
  const server = await ensureLoopbackServers(port);
  pendingCallbacks.set(state, { resolve: resolveCode, reject: rejectCode, successRedirectUrl, consumed: false });
  let timer;
  try {
    const redirectUri = `http://localhost:${server.port}`;
    timer = setTimeout(() => rejectCode(new Error('INTERACTIVE_LOGIN_TIMEOUT')), timeoutMs);
    const authorizationUrl = await pca.getAuthCodeUrl({ scopes, redirectUri, state, responseMode: 'query', codeChallenge: challenge, codeChallengeMethod: 'S256' });
    onReady({ authorizationUrl });
    const code = await result;
    return await pca.acquireTokenByCode({ scopes, redirectUri, code, codeVerifier: verifier });
  } finally {
    clearTimeout(timer);
    pendingCallbacks.delete(state);
  }
}
