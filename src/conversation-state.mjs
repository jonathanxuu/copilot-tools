import { existsSync, mkdirSync, readFileSync, chmodSync, writeFileSync, renameSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

const emptyState = () => ({ conversations: {} });

export function createConversationState(filePath) {
  const path = resolve(filePath);
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  let state = emptyState();
  if (existsSync(path)) {
    try {
      const parsed = JSON.parse(readFileSync(path, 'utf8'));
      if (parsed && typeof parsed === 'object' && parsed.conversations && typeof parsed.conversations === 'object') state = parsed;
    } catch {
      // A corrupt state file must not prevent the MCP server from starting.
    }
  }
  try { chmodSync(path, 0o600); } catch {}
  const save = () => {
    const temporary = `${path}.${process.pid}.tmp`;
    writeFileSync(temporary, JSON.stringify(state, null, 2), { encoding: 'utf8', mode: 0o600 });
    chmodSync(temporary, 0o600);
    renameSync(temporary, path);
  };
  return {
    get(userId, handle) {
      const item = state.conversations[`${userId}:${handle}`];
      return item?.remoteConversationId ?? null;
    },
    set(userId, handle, remoteConversationId) {
      state.conversations[`${userId}:${handle}`] = { remoteConversationId, updatedAt: new Date().toISOString() };
      save();
    },
    removeUser(userId) {
      const prefix = `${userId}:`;
      for (const key of Object.keys(state.conversations)) if (key.startsWith(prefix)) delete state.conversations[key];
      save();
    },
  };
}
