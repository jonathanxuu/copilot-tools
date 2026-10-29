import { createHash, randomBytes } from 'node:crypto';

const pending = new Map();

const makeState = () => randomBytes(32).toString('base64url');
const makeVerifier = () => randomBytes(32).toString('base64url');

function challenge(verifier) {
  return createHash('sha256').update(verifier).digest('base64url');
}

function html(title, message) {
  return `<!doctype html><html lang="en"><meta charset="utf-8"><title>${title}</title><p>${message}</p>`;
}

export function installBrowserCallback(app, { path = '/microsoft/callback' } = {}) {
  app.get(path, async (req, res) => {
    const state = typeof req.query.state === 'string' ? req.query.state : '';
    const item = pending.get(state);
    if (!item || item.expiresAt < Date.now() || item.consumed) {
      res.status(400).type('html').send(html('Microsoft login', 'This Microsoft login has expired. Return to Counso and connect again.'));
      return;
    }
    item.consumed = true;
    pending.delete(state);
    if (req.query.error) {
      const detail = typeof req.query.error_description === 'string' ? req.query.error_description : String(req.query.error);
      item.reject(new Error(`MICROSOFT_INTERACTIVE_DENIED:${detail.slice(0, 300)}`));
      res.status(400).type('html').send(html('Microsoft login', 'Microsoft did not complete the login. Return to Counso and try again.'));
      return;
    }
    const code = typeof req.query.code === 'string' ? req.query.code : '';
    if (!code) {
      item.reject(new Error('MISSING_AUTHORIZATION_CODE'));
      res.status(400).type('html').send(html('Microsoft login', 'The Microsoft authorization code was missing. Return to Counso and try again.'));
      return;
    }
    try {
      const result = await item.pca.acquireTokenByCode({
        scopes: item.scopes,
        redirectUri: item.redirectUri,
        code,
        codeVerifier: item.verifier,
      });
      item.resolve(result);
      res.redirect(item.successRedirectUrl);
    } catch (error) {
      console.warn(JSON.stringify({
        event: 'microsoft_browser_callback_failed',
        errorCode: error?.errorCode ?? null,
        subError: error?.subError ?? null,
        correlationId: error?.correlationId ?? null,
        message: error instanceof Error ? error.message.slice(0, 500) : 'MICROSOFT_TOKEN_EXCHANGE_FAILED',
      }));
      item.reject(error);
      res.status(400).type('html').send(html('Microsoft login', 'Microsoft login could not be completed. Return to Counso and try again.'));
    }
  });
}

export async function browserLogin(pca, scopes, onReady, { redirectUri, successRedirectUrl, timeoutMs = 300000 } = {}) {
  if (!redirectUri || !/^https:\/\//i.test(redirectUri)) throw new Error('CONFIG_BROWSER_REDIRECT_HTTPS_REQUIRED');
  const state = makeState();
  const verifier = makeVerifier();
  const authorizationUrl = await pca.getAuthCodeUrl({
    scopes,
    redirectUri,
    state,
    responseMode: 'query',
    codeChallenge: challenge(verifier),
    codeChallengeMethod: 'S256',
  });
  let timer;
  const result = new Promise((resolve, reject) => {
    pending.set(state, { pca, scopes, redirectUri, verifier, successRedirectUrl, resolve, reject, consumed: false, expiresAt: Date.now() + timeoutMs });
    timer = setTimeout(() => {
      pending.delete(state);
      reject(new Error('INTERACTIVE_LOGIN_TIMEOUT'));
    }, timeoutMs);
  });
  void result.catch(() => {});
  onReady({ authorizationUrl, mode: 'browser' });
  try { return await result; } finally { clearTimeout(timer); pending.delete(state); }
}
