import { createCipheriv, createDecipheriv, createHash, randomBytes } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

function key() {
  const raw = process.env.M365_MCP_TOKEN_CACHE_KEY;
  if (!raw) throw new Error('M365_MCP_TOKEN_CACHE_KEY_REQUIRED');
  return /^[0-9a-f]{64}$/i.test(raw) ? Buffer.from(raw, 'hex') : createHash('sha256').update(raw).digest();
}

function path() { return process.env.M365_MCP_TOKEN_CACHE_PATH ?? './data/msal-cache.enc'; }

export function createMsalCachePlugin() {
  return {
    async beforeCacheAccess(cacheContext) {
      try {
        const encoded = JSON.parse(await readFile(path(), 'utf8'));
        const decipher = createDecipheriv('aes-256-gcm', key(), Buffer.from(encoded.iv, 'base64'));
        decipher.setAuthTag(Buffer.from(encoded.tag, 'base64'));
        const data = Buffer.concat([decipher.update(Buffer.from(encoded.data, 'base64')), decipher.final()]).toString('utf8');
        cacheContext.tokenCache.deserialize(data);
      } catch (error) {
        if (error?.code !== 'ENOENT') throw new Error('M365_MCP_TOKEN_CACHE_READ_FAILED');
      }
    },
    async afterCacheAccess(cacheContext) {
      if (!cacheContext.cacheHasChanged) return;
      const iv = randomBytes(12);
      const cipher = createCipheriv('aes-256-gcm', key(), iv);
      const data = Buffer.concat([cipher.update(cacheContext.tokenCache.serialize(), 'utf8'), cipher.final()]);
      await mkdir(dirname(path()), { recursive: true });
      await writeFile(path(), JSON.stringify({ iv: iv.toString('base64'), tag: cipher.getAuthTag().toString('base64'), data: data.toString('base64') }), { mode: 0o600 });
    },
  };
}
