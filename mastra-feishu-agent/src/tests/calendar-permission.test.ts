import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { CardActionEvent, LarkChannel } from '@larksuiteoapi/node-sdk';

process.env.AGENT_DATA_DIR = mkdtempSync(join(tmpdir(), 'mastra-calendar-test-'));
process.env.FEISHU_ALLOWED_USERS = 'u1';
process.env.FEISHU_ALLOWED_CHATS = 'dm1';
process.env.FEISHU_APP_ID = 'cli_test';
process.env.FEISHU_APP_SECRET = 'test-secret';
process.env.FEISHU_DOMAIN = 'https://api.test';
process.env.FEISHU_OAUTH_AUTHORIZE_URL = 'https://accounts.test/open-apis/authen/v1/authorize';
process.env.FEISHU_OAUTH_TOKEN_URL = 'https://api.test/open-apis/authen/v2/oauth/token';
process.env.FEISHU_OAUTH_REDIRECT_URI = 'http://127.0.0.1:4312/oauth/callback';
process.env.FEISHU_TOKEN_ENCRYPTION_KEY = 'test-only-random-encryption-passphrase';

const caller = { userId: 'u1', chatId: 'dm1', chatType: 'p2p' as const };
let permission: typeof import('../office/permission-choice');
let bridge: typeof import('../runtime/approval-bridge');
let officeBridge: typeof import('../runtime/office-bridge');
let store: typeof import('../state/store');
let auth: typeof import('../office/user-auth');

beforeAll(async () => {
  permission = await import('../office/permission-choice');
  bridge = await import('../runtime/approval-bridge');
  officeBridge = await import('../runtime/office-bridge');
  store = await import('../state/store');
  auth = await import('../office/user-auth');
});
afterEach(async () => {
  await (await store.stateDb()).execute('DELETE FROM user_tokens');
  await (await store.stateDb()).execute('DELETE FROM permission_requests');
  vi.unstubAllGlobals();
  permission.setPermissionNotifier(undefined);
  bridge.setApprovalNotifier(undefined);
  officeBridge.setOfficeApi(undefined);
});

it('binds OAuth to the requester, then requests write approval with the user calendar', async () => {
  const questions: string[] = [];
  const cards: Array<{ id: string; summary: string }> = [];
  const requests: Array<{ url: string; userAccessToken?: string }> = [];
  permission.setPermissionNotifier(async (_chat, text) => { questions.push(text); });
  bridge.setApprovalNotifier(async (_chat, id, summary) => { cards.push({ id, summary }); });
  const { OfficeApi } = await import('../office/client');
  officeBridge.setOfficeApi(new OfficeApi({
    request: async request => {
      requests.push(request);
      if (request.url.endsWith('/primary')) return { code: 0, data: { calendars: [{ calendar: { calendar_id: 'my-calendar', type: 'primary', role: 'owner' } }] } };
      return { code: 0, data: {} };
    }, send: async () => undefined,
  }));
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => {
    const url = String(input);
    if (url.includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'u-test-token', refresh_token: 'r-test-token', expires_in: 3600 }), { status: 200 });
    return new Response(JSON.stringify({ code: 0, data: { open_id: 'u1' } }), { status: 200 });
  }));

  const result = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'user', summary: 'Meeting', startTime: '1760000000', endTime: '1760003600' });
  expect('requestId' in result).toBe(true);
  expect(cards).toHaveLength(0);
  expect(await permission.resolvePermissionReply(caller, '用户权限')).toBe(true);
  expect(cards).toHaveLength(0);
  const link = questions.at(-1)!.match(/https:\/\/\S+/)![0];
  const state = new URL(link).searchParams.get('state')!;
  expect(state).toBe('requestId' in result ? result.requestId : '');
  await permission.completeCalendarAuthorization(state, 'one-time-code');
  expect(cards).toHaveLength(1);
  expect(cards[0].summary).toContain('用户身份');
  expect(requests).toMatchObject([{ url: '/open-apis/calendar/v4/calendars/primary', userAccessToken: 'u-test-token' }]);
  expect(await auth.userAccessToken('u1')).toBe('u-test-token');
  const row = await store.getAction(cards[0].id);
  expect(JSON.parse(row!.payload)).toMatchObject({ permissionMode: 'user', calendarId: 'my-calendar' });
  await expect(permission.completeCalendarAuthorization(state, 'one-time-code')).rejects.toThrow('已失效');

  officeBridge.setOfficeApi(new OfficeApi({ request: async () => {
    throw Object.assign(new Error('Request failed with status code 400'), {
      response: { status: 400, data: { code: 99991679, msg: 'calendar read scope required' } },
    });
  }, send: async () => undefined }));
  const second = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'user', summary: 'Second', startTime: '1760000000', endTime: '1760003600' });
  expect('requestId' in second).toBe(true);
  expect(cards).toHaveLength(1);
  expect(questions.at(-1)).toContain('重新授权');
  await permission.resolvePermissionReply(caller, '用户权限');
  expect(cards).toHaveLength(1);
  expect(questions.at(-1)).toContain('https://accounts.test/');
});

it('rejects an OAuth callback from a different user before storing a token or sending approval', async () => {
  const questions: string[] = [];
  const cards: string[] = [];
  permission.setPermissionNotifier(async (_chat, text) => { questions.push(text); });
  bridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => {
    if (String(input).includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'wrong-user-token', expires_in: 3600 }), { status: 200 });
    return new Response(JSON.stringify({ code: 0, data: { open_id: 'u2' } }), { status: 200 });
  }));
  const result = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'user', summary: 'Private', startTime: '1760000000', endTime: '1760003600' });
  if (!('requestId' in result)) throw new Error('Expected permission choice');
  await permission.resolvePermissionReply(caller, '用户权限');
  await expect(permission.completeCalendarAuthorization(result.requestId, 'wrong-user-code')).rejects.toThrow('不一致');
  expect(cards).toHaveLength(0);
  expect(await auth.userAccessToken('u1')).toBe(null);
});

it('requests calendar read permission as well as write permission for primary-calendar lookup', async () => {
  const scope = new URL(auth.authorizationUrl('test-state')).searchParams.get('scope')?.split(' ') ?? [];
  expect(scope).toContain('calendar:calendar');
  expect(scope).toContain('calendar:calendar:readonly');
});

it('restarts authorization after a failed callback without reusing its token or failure state', async () => {
  const messages: string[] = [];
  const cards: string[] = [];
  permission.setPermissionNotifier(async (_chat, text) => { messages.push(text); });
  bridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
  const { OfficeApi } = await import('../office/client');
  officeBridge.setOfficeApi(new OfficeApi({
    request: async () => { throw Object.assign(new Error('Request failed with status code 400'), {
      response: { status: 400, data: { code: 99991679, msg: 'calendar read scope required' } },
    }); }, send: async () => undefined,
  }));
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => {
    if (String(input).includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'under-scoped-token', refresh_token: 'refresh-token', expires_in: 3600 }), { status: 200 });
    return new Response(JSON.stringify({ code: 0, data: { open_id: 'u1' } }), { status: 200 });
  }));

  const first = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'user', summary: 'First', startTime: '1760000000', endTime: '1760003600' });
  if (!('requestId' in first)) throw new Error('Expected permission request');
  await permission.resolvePermissionReply(caller, '用户权限');
  await expect(permission.completeCalendarAuthorization(first.requestId, 'code-1')).rejects.toThrow('99991679');
  expect(await auth.userAccessToken(caller.userId)).toBe(null);

  const second = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'auto', summary: 'Second', startTime: '1760000000', endTime: '1760003600' });
  if (!('requestId' in second)) throw new Error('Expected new permission request');
  expect(second.requestId).not.toBe(first.requestId);
  const old = await (await store.stateDb()).execute({ sql: 'SELECT status FROM permission_requests WHERE id = ?', args: [first.requestId] });
  expect(old.rows[0]?.status).toBe('superseded');
  expect(await permission.resolvePermissionReply(caller, '用户权限')).toBe(true);
  expect(messages.at(-1)).toContain(`state=${second.requestId}`);
  expect(cards).toHaveLength(0);
});

it('accepts a direct reauthorization reply to a recent failed calendar request', async () => {
  const messages: string[] = [];
  permission.setPermissionNotifier(async (_chat, text) => { messages.push(text); });
  bridge.setApprovalNotifier(async () => undefined);
  const { OfficeApi } = await import('../office/client');
  officeBridge.setOfficeApi(new OfficeApi({
    request: async () => { throw Object.assign(new Error('Request failed with status code 400'), {
      response: { status: 400, data: { code: 99991679, msg: 'calendar read scope required' } },
    }); }, send: async () => undefined,
  }));
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => {
    if (String(input).includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'under-scoped-token', expires_in: 3600 }), { status: 200 });
    return new Response(JSON.stringify({ code: 0, data: { open_id: 'u1' } }), { status: 200 });
  }));
  const first = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'user', summary: 'Meeting', startTime: '1760000000', endTime: '1760003600' });
  if (!('requestId' in first)) throw new Error('Expected permission request');
  await permission.resolvePermissionReply(caller, '用户权限');
  await expect(permission.completeCalendarAuthorization(first.requestId, 'code-1')).rejects.toThrow('99991679');
  expect(messages.at(-1)).toContain('重新授权');

  expect(await permission.resolvePermissionReply(caller, '重新授权')).toBe(true);
  const url = messages.at(-1)!.match(/https:\/\/\S+/)?.[0];
  expect(url).toBeTruthy();
  const nextId = new URL(url!).searchParams.get('state');
  expect(nextId).not.toBe(first.requestId);
  const rows = await (await store.stateDb()).execute({ sql: 'SELECT id, status FROM permission_requests ORDER BY created_at DESC' });
  expect(rows.rows.find(row => row.id === first.requestId)?.status).toBe('superseded');
  expect(rows.rows.find(row => row.id === nextId)?.status).toBe('auth');
});

it('turns an existing token permission failure into a fresh OAuth request after an auto choice', async () => {
  const messages: string[] = [];
  permission.setPermissionNotifier(async (_chat, text) => { messages.push(text); });
  bridge.setApprovalNotifier(async () => undefined);
  const { OfficeApi } = await import('../office/client');
  officeBridge.setOfficeApi(new OfficeApi({
    request: async () => ({ code: 0, data: { calendars: [{ calendar: { calendar_id: 'my-calendar', type: 'primary', role: 'owner' } }] } }),
    send: async () => undefined,
  }));
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => {
    if (String(input).includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'old-token', refresh_token: 'refresh-token', expires_in: 3600 }), { status: 200 });
    return new Response(JSON.stringify({ code: 0, data: { open_id: 'u1' } }), { status: 200 });
  }));
  const initial = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'user', summary: 'Initial', startTime: '1760000000', endTime: '1760003600' });
  if (!('requestId' in initial)) throw new Error('Expected permission request');
  await permission.resolvePermissionReply(caller, '用户权限');
  await permission.completeCalendarAuthorization(initial.requestId, 'code-1');

  officeBridge.setOfficeApi(new OfficeApi({
    request: async () => { throw Object.assign(new Error('Request failed with status code 400'), {
      response: { status: 400, data: { code: 99991679, msg: 'calendar read scope required' } },
    }); }, send: async () => undefined,
  }));
  const next = await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'auto', summary: 'Next', startTime: '1760000000', endTime: '1760003600' });
  if (!('requestId' in next)) throw new Error('Expected permission request');
  expect(await permission.resolvePermissionReply(caller, '用户权限')).toBe(true);
  expect(messages.at(-1)).toContain(`state=${next.requestId}`);
  expect(await auth.userAccessToken(caller.userId)).toBe(null);
});

it('sends only the permission question when the agent later summarizes stale authorization history', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { memoryKeys } = await import('../identity');
  const sent: string[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chat: string, content: { text?: string; markdown?: string }) => { sent.push(content.text ?? content.markdown ?? ''); return {}; },
  } as unknown as LarkChannel);
  permission.setPermissionNotifier(async (chat, text) => { await gateway.channel.send(chat, { text }); });
  let subscriber: (event: unknown) => void = () => undefined;
  const session = {
    subscribe: (callback: (event: unknown) => void) => { subscriber = callback; return () => undefined; },
    sendMessage: async () => {
      await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'auto', summary: 'Meeting', startTime: '1760000000', endTime: '1760003600' });
      subscriber({ type: 'message_update', event: { type: 'text-delta', delta: '旧记录显示创建失败，请提供 calendarId。' } });
      subscriber({ type: 'agent_end', reason: 'completed' });
    },
  };
  Reflect.get(gateway, 'sessions').set(memoryKeys(caller).sessionId, session);
  await Reflect.get(gateway, 'runAgent').call(gateway, '创建日程', caller, 'message-1');
  await vi.waitFor(() => expect(Reflect.get(gateway, 'activeRuns').size).toBe(0));
  expect(sent).toHaveLength(1);
  expect(sent[0]).toContain('用户权限');
});

it('does not send an old agent summary when the requester cancels a new permission question immediately', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { memoryKeys } = await import('../identity');
  const sent: string[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chat: string, content: { text?: string; markdown?: string }) => { sent.push(content.text ?? content.markdown ?? ''); return {}; },
  } as unknown as LarkChannel);
  permission.setPermissionNotifier(async (chat, text) => { await gateway.channel.send(chat, { text }); });
  let subscriber: (event: unknown) => void = () => undefined;
  Reflect.get(gateway, 'sessions').set(memoryKeys(caller).sessionId, {
    subscribe: (callback: (event: unknown) => void) => { subscriber = callback; return () => undefined; },
    sendMessage: async () => {
      await permission.requestCalendarPermission(caller, { kind: 'calendar_create', permissionMode: 'auto', summary: 'Meeting', startTime: '1760000000', endTime: '1760003600' });
      await permission.resolvePermissionReply(caller, '取消');
      subscriber({ type: 'message_update', event: { type: 'text-delta', delta: '旧记录显示创建失败。' } });
      subscriber({ type: 'agent_end', reason: 'completed' });
    },
  });
  await Reflect.get(gateway, 'runAgent').call(gateway, '创建日程', caller, 'message-2');
  await vi.waitFor(() => expect(Reflect.get(gateway, 'activeRuns').size).toBe(0));
  expect(sent).toHaveLength(2);
  expect(sent[1]).toContain('已取消');
});

it('persists the user-requested tomorrow date from trusted Feishu context before asking for authorization', async () => {
  const { createCalendarTool } = await import('../mastra/tools/office-tools');
  permission.setPermissionNotifier(async () => undefined);
  const requestContext = { get: (key: string) => ({
    caller,
    originalMessage: '请为我创建一个明天上午9点到10点的会议日程，主题是10月工作总结大会',
    messageCreatedAt: Date.parse('2026-10-26T10:55:00+08:00'),
  })[key] };
  const execute = Reflect.get(createCalendarTool, 'execute') as (input: unknown, context: unknown) => Promise<{ requestId: string }>;
  const result = await execute({ permissionMode: 'auto', summary: '10月工作总结大会', startTime: '1792717200', endTime: '1792720800' }, { requestContext });
  const stored = await (await store.stateDb()).execute({ sql: 'SELECT payload FROM permission_requests WHERE id = ?', args: [result.requestId] });
  expect(JSON.parse(String(stored.rows[0]?.payload))).toMatchObject({ startTime: '1793062800', endTime: '1793066400' });
});
