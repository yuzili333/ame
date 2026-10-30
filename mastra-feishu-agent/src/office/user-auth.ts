import { createCipheriv, createDecipheriv, randomBytes, scryptSync } from 'node:crypto';
import { createServer, type Server } from 'node:http';
import { config } from '../config';
import { stateDb } from '../state/store';

type OAuthResponse = { code?: number | string; msg?: string; access_token?: string; refresh_token?: string; expires_in?: number; scope?: string };
export const CALENDAR_SCOPES = ['calendar:calendar', 'calendar:calendar:readonly', 'offline_access'];
const required = () => {
  if (!config.appId || !config.appSecret || !config.domain || !config.oauthAuthorizeUrl || !config.oauthTokenUrl || !config.oauthRedirectUri || !config.tokenEncryptionKey) {
    throw new Error('用户权限尚未配置：需设置 FEISHU_OAUTH_AUTHORIZE_URL、FEISHU_OAUTH_TOKEN_URL、FEISHU_OAUTH_REDIRECT_URI 和 FEISHU_TOKEN_ENCRYPTION_KEY');
  }
};

function key() {
  required();
  return scryptSync(config.tokenEncryptionKey!, 'mastra-feishu-user-token-v1', 32);
}

function encrypt(value: string) {
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', key(), iv);
  const body = Buffer.concat([cipher.update(value, 'utf8'), cipher.final()]);
  return Buffer.concat([iv, cipher.getAuthTag(), body]).toString('base64');
}

function decrypt(value: string) {
  const bytes = Buffer.from(value, 'base64');
  const decipher = createDecipheriv('aes-256-gcm', key(), bytes.subarray(0, 12));
  decipher.setAuthTag(bytes.subarray(12, 28));
  return Buffer.concat([decipher.update(bytes.subarray(28)), decipher.final()]).toString('utf8');
}

export function authorizationUrl(state: string, scopes: string[] = CALENDAR_SCOPES): string {
  required();
  const url = new URL(config.oauthAuthorizeUrl!);
  url.searchParams.set('client_id', config.appId!);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('redirect_uri', config.oauthRedirectUri!);
  url.searchParams.set('scope', [...new Set(scopes)].join(' '));
  url.searchParams.set('state', state);
  return url.toString();
}

async function tokenRequest(body: Record<string, string>): Promise<OAuthResponse> {
  required();
  const response = await fetch(config.oauthTokenUrl!, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...body, client_id: config.appId, client_secret: config.appSecret }),
    signal: AbortSignal.timeout(15_000),
  });
  const data = await response.json() as OAuthResponse;
  if (!response.ok || Number(data.code ?? 0) !== 0 || !data.access_token || !data.expires_in) {
    throw new Error(`飞书用户授权失败：${data.msg || response.status}`);
  }
  return data;
}

async function saveToken(userId: string, data: OAuthResponse, scopes: string[]) {
  const db = await stateDb();
  const granted = data.scope?.split(/\s+/).filter(Boolean) ?? scopes;
  await db.execute({ sql: `INSERT INTO user_tokens(user_id, encrypted_token, encrypted_refresh_token, expires_at, updated_at, scopes)
    VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT(user_id) DO UPDATE SET encrypted_token=excluded.encrypted_token,
    encrypted_refresh_token=excluded.encrypted_refresh_token, expires_at=excluded.expires_at, updated_at=excluded.updated_at, scopes=excluded.scopes`,
    args: [userId, encrypt(data.access_token!), data.refresh_token ? encrypt(data.refresh_token) : null, Date.now() + data.expires_in! * 1000, Date.now(), JSON.stringify(granted)] });
}

export async function exchangeCodeForUser(code: string, expectedUserId: string, scopes: string[] = CALENDAR_SCOPES): Promise<string> {
  const data = await tokenRequest({ grant_type: 'authorization_code', code, redirect_uri: config.oauthRedirectUri! });
  const response = await fetch(new URL('/open-apis/authen/v1/user_info', config.domain).toString(), {
    headers: { Authorization: `Bearer ${data.access_token}` }, signal: AbortSignal.timeout(15_000),
  });
  const identity = await response.json() as { code?: number; msg?: string; data?: { open_id?: string } };
  if (!response.ok || identity.code !== 0 || identity.data?.open_id !== expectedUserId) {
    throw new Error('授权账号与原请求者不一致，未保存用户令牌');
  }
  await saveToken(expectedUserId, data, scopes);
  return data.access_token!;
}

export async function invalidateUserAccessToken(userId: string, expectedToken: string): Promise<void> {
  const db = await stateDb();
  const row = (await db.execute({ sql: 'SELECT encrypted_token FROM user_tokens WHERE user_id = ?', args: [userId] })).rows[0];
  if (!row || decrypt(String(row.encrypted_token)) !== expectedToken) return;
  await db.execute({ sql: 'DELETE FROM user_tokens WHERE user_id = ? AND encrypted_token = ?', args: [userId, String(row.encrypted_token)] });
}

const refreshing = new Map<string, Promise<string>>();
export async function userAccessToken(userId: string, requiredScopes: string[] = []): Promise<string | null> {
  const db = await stateDb();
  const row = (await db.execute({ sql: 'SELECT * FROM user_tokens WHERE user_id = ?', args: [userId] })).rows[0];
  if (!row) return null;
  const scopes = row.scopes ? JSON.parse(String(row.scopes)) as string[] : CALENDAR_SCOPES;
  if (!requiredScopes.filter(scope => scope !== 'offline_access').every(scope => scopes.includes(scope))) return null;
  if (Number(row.expires_at) > Date.now() + 60_000) return decrypt(String(row.encrypted_token));
  if (!row.encrypted_refresh_token) return null;
  let pending = refreshing.get(userId);
  if (!pending) {
    pending = (async () => {
      try {
        const next = await tokenRequest({ grant_type: 'refresh_token', refresh_token: decrypt(String(row.encrypted_refresh_token)) });
        if (!next.refresh_token) next.refresh_token = decrypt(String(row.encrypted_refresh_token));
        await saveToken(userId, next, scopes);
        return next.access_token!;
      } catch {
        const current = (await db.execute({ sql: 'SELECT encrypted_token, expires_at FROM user_tokens WHERE user_id = ?', args: [userId] })).rows[0];
        if (current && String(current.encrypted_token) !== String(row.encrypted_token) && Number(current.expires_at) > Date.now() + 60_000) {
          return decrypt(String(current.encrypted_token));
        }
        await db.execute({ sql: 'DELETE FROM user_tokens WHERE user_id = ? AND encrypted_token = ?', args: [userId, String(row.encrypted_token)] });
        throw new Error('用户日历授权已过期，请重新授权');
      }
    })().finally(() => refreshing.delete(userId));
    refreshing.set(userId, pending);
  }
  return pending;
}

export async function startOAuthCallback(onAuthorized: (state: string, code: string) => Promise<void>): Promise<Server | undefined> {
  if (!config.oauthRedirectUri) return undefined;
  const redirect = new URL(config.oauthRedirectUri);
  const server = createServer((request, response) => {
    void (async () => {
      const url = new URL(request.url ?? '/', 'http://localhost');
      if (request.method !== 'GET' || url.pathname !== redirect.pathname) { response.writeHead(404).end(); return; }
      const state = url.searchParams.get('state');
      const code = url.searchParams.get('code');
      if (!state || !code || url.searchParams.has('error')) { response.writeHead(400).end('授权已取消或参数不完整'); return; }
      try {
        await onAuthorized(state, code);
        response.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' }).end('飞书用户授权成功，请返回飞书查看后续操作。');
      } catch (error) {
        response.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' }).end(error instanceof Error ? error.message : '授权失败');
      }
    })();
  });
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(config.oauthListenPort, config.oauthListenHost, () => {
      server.off('error', reject);
      resolve();
    });
  });
  return server;
}
