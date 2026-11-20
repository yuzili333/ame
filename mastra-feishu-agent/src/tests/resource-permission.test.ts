import { afterEach, beforeAll, beforeEach, expect, it, vi } from 'vitest';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

process.env.AGENT_DATA_DIR = mkdtempSync(join(tmpdir(), 'resource-auth-'));
process.env.FEISHU_ALLOWED_USERS = 'u1';
process.env.FEISHU_ALLOWED_CHATS = 'dm1';
process.env.FEISHU_APP_ID = 'test';
process.env.FEISHU_APP_SECRET = 'secret';
process.env.FEISHU_DOMAIN = 'https://api.test';
process.env.FEISHU_OAUTH_AUTHORIZE_URL = 'https://accounts.test/authorize';
process.env.FEISHU_OAUTH_TOKEN_URL = 'https://api.test/oauth/token';
process.env.FEISHU_OAUTH_REDIRECT_URI = 'http://localhost:4312/callback';
process.env.FEISHU_TOKEN_ENCRYPTION_KEY = 'test-only-key';

const caller = { userId: 'u1', chatId: 'dm1', chatType: 'p2p' as const, messageId: 'origin' };
let permission: typeof import('../office/resource-permission');
let store: typeof import('../state/store');
let auth: typeof import('../office/user-auth');
let bridge: typeof import('../runtime/office-bridge');
let approvals: typeof import('../state/approvals');
const cards: string[] = [];
const messages: string[] = [];
const requests: import('../office/client').ApiRequest[] = [];
let responseCode = 0;
let role = 'owner';
let calendarType = 'shared';
let account = 'u1';

beforeAll(async () => {
  permission = await import('../office/resource-permission');
  store = await import('../state/store');
  auth = await import('../office/user-auth');
  bridge = await import('../runtime/office-bridge');
  approvals = await import('../state/approvals');
});
beforeEach(async () => {
  cards.length = messages.length = requests.length = 0;
  responseCode = 0; role = 'owner'; calendarType = 'shared'; account = 'u1';
  const { setApprovalNotifier } = await import('../runtime/approval-bridge');
  setApprovalNotifier(async (_chat, id) => { cards.push(id); });
  bridge.setOfficeTextSender(async (_chat, text) => { messages.push(text); });
  const { OfficeApi } = await import('../office/client');
  bridge.setOfficeApi(new OfficeApi({ request: async request => {
    requests.push(request);
    return { code: responseCode, msg: 'permission denied', data: { calendar: { calendar_id: 'cal', type: calendarType, role }, items: [{ guid: 't1', summary: 'My task' }], has_more: true, page_token: 'next' } };
  }, send: async () => undefined }));
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => new Response(JSON.stringify(String(input).includes('/oauth/token')
    ? { code: 0, access_token: 'user-token', refresh_token: 'refresh', expires_in: 3600 }
    : { code: 0, data: { open_id: account } }), { status: 200 })));
});
afterEach(async () => {
  const db = await store.stateDb();
  for (const table of ['user_tokens', 'permission_requests', 'actions', 'recalled_messages', 'queued_messages']) await db.execute(`DELETE FROM ${table}`);
  vi.unstubAllGlobals();
  permission.setResourceContinuation(undefined);
});

it('continues the original natural-language request when OAuth interrupted a prerequisite lookup', async () => {
  const resumed: unknown[] = [];
  permission.setResourceContinuation(async (...args) => { resumed.push(args); });
  await permission.requestResourceOperation(caller, { kind: 'task_list' }, { originalMessage: '把我负责的“周报”任务标为完成', messageCreatedAt: 1234 });
  await permission.completeResourceAuthorization(await consent(), 'code');
  expect(resumed).toHaveLength(1);
  expect(resumed[0]).toMatchObject([caller, { originalMessage: '把我负责的“周报”任务标为完成', messageCreatedAt: 1234 }, { items: [{ guid: 't1' }] }, expect.any(String)]);
});

it('queues authorized lookup continuation while the chat is busy and restores the original caller and time', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const gateway = new FeishuGateway({ rawClient: {} } as unknown as import('@larksuiteoapi/node-sdk').LarkChannel);
  Reflect.get(gateway, 'activeRuns').set('busy', { chatId: caller.chatId });
  const run = vi.fn(async () => undefined);
  Reflect.set(gateway, 'runAgent', run);
  permission.setResourceContinuation((...args) => Reflect.get(gateway, 'queueResourceContinuation').apply(gateway, args));
  await permission.requestResourceOperation(caller, { kind: 'task_list' }, { originalMessage: '把周报任务标为完成', messageCreatedAt: 1234 });
  const state = await consent();
  await permission.completeResourceAuthorization(state, 'code');
  const queued = (await (await store.stateDb()).execute('SELECT * FROM queued_messages')).rows;
  expect(queued).toHaveLength(1);
  expect(run).not.toHaveBeenCalled();
  Reflect.get(gateway, 'activeRuns').clear();
  await Reflect.get(gateway, 'drainQueue').call(gateway, caller.chatId);
  expect(run).toHaveBeenCalledWith('把周报任务标为完成', caller, caller.messageId, 1234, false, expect.stringContaining('My task'));
  await Reflect.get(gateway, 'drainQueue').call(gateway, caller.chatId);
  expect(run).toHaveBeenCalledTimes(1);
});

async function consent() {
  const decision = await approvals.decideAction(cards[0], caller, true, (action, key) => {
    if (action.kind !== 'resource_authorize') throw new Error('Expected authorization');
    return permission.beginResourceAuthorization(action, key);
  });
  expect(decision.status).toBe('succeeded');
  return new URL((decision.result as { authorizationUrl: string }).authorizationUrl).searchParams.get('state')!;
}

it('requests approval before revealing OAuth, then resumes task list exactly once with the user token', async () => {
  const result = await permission.requestResourceOperation(caller, { kind: 'task_list', pageSize: 20, completed: false });
  expect(result.status).toBe('authorization_required');
  expect(cards).toHaveLength(1);
  expect(messages.join('')).not.toContain('https://');
  expect(requests).toEqual([]);
  const state = await consent();
  await permission.completeResourceAuthorization(state, 'code');
  expect(requests).toMatchObject([{ method: 'GET', url: '/open-apis/task/v2/tasks', userAccessToken: 'user-token', params: { page_size: 20, completed: false } }]);
  expect(messages.join('')).toContain('My task');
  expect(messages.join('')).toContain('next');
  await expect(permission.completeResourceAuthorization(state, 'code')).rejects.toThrow('失效');
  expect(requests).toHaveLength(1);
});

it('resumes deletion with an operation approval; OAuth itself does not delete', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_delete', taskId: 't1' });
  await permission.completeResourceAuthorization(await consent(), 'code');
  expect(cards).toHaveLength(2);
  const stored = await store.getAction(cards[1]);
  expect(JSON.parse(stored!.payload)).toMatchObject({ kind: 'task_delete', taskId: 't1', permissionMode: 'user' });
  expect(requests.every(request => request.method === 'GET')).toBe(true);
});

it('does not allow another requester to approve or authorize', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_list' });
  const execute = vi.fn();
  expect((await approvals.decideAction(cards[0], { ...caller, userId: 'u2' }, true, execute)).status).toBe('denied');
  expect(execute).not.toHaveBeenCalled();
  const state = await consent(); account = 'u2';
  await expect(permission.completeResourceAuthorization(state, 'code')).rejects.toThrow('不一致');
  expect(await auth.userAccessToken('u1')).toBeNull();
  expect(requests).toHaveLength(0);
});

it('requests user authorization when an application cannot access the selected shared calendar', async () => {
  responseCode = 99991672;
  const result = await permission.requestResourceOperation(caller, { kind: 'calendar_read', calendarId: 'cal', permissionMode: 'app' });
  expect(result.status).toBe('authorization_required');
  const state = await consent(); responseCode = 0;
  await permission.completeResourceAuthorization(state, 'code');
  expect(requests.at(-1)).toMatchObject({ url: '/open-apis/calendar/v4/calendars/cal', userAccessToken: 'user-token' });
});

it('does not try OAuth again when the authorized user only has reader access', async () => {
  await auth.exchangeCodeForUser('code', 'u1', auth.CALENDAR_SCOPES);
  role = 'reader';
  await expect(permission.requestResourceOperation(caller, { kind: 'calendar_delete', calendarId: 'cal', eventId: 'evt', permissionMode: 'user' })).rejects.toThrow('所有者');
  expect(cards).toHaveLength(0);
});

it('rejects deletion of a personal primary calendar before operation approval', async () => {
  await auth.exchangeCodeForUser('code', 'u1', auth.CALENDAR_SCOPES); calendarType = 'primary';
  await expect(permission.requestResourceOperation(caller, { kind: 'calendar_container_delete', calendarId: 'cal' })).rejects.toThrow('主日历');
  expect(cards).toHaveLength(0);
});

it('cancels pending authorization and never resumes it', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_list' });
  const state = await consent();
  expect(await permission.resolveResourcePermissionReply(caller, '取消')).toBe(true);
  await expect(permission.completeResourceAuthorization(state, 'code')).rejects.toThrow('失效');
  expect(requests).toHaveLength(0);
});

it('requires task scope even when an existing calendar token is valid', async () => {
  await auth.exchangeCodeForUser('code', 'u1', auth.CALENDAR_SCOPES);
  await permission.requestResourceOperation(caller, { kind: 'task_list' });
  expect(cards).toHaveLength(1);
  const state = await consent();
  const row = (await (await store.stateDb()).execute({ sql: 'SELECT payload FROM permission_requests WHERE id=?', args: [state] })).rows[0];
  expect(String(row.payload)).toContain('calendar:calendar');
  expect(String(row.payload)).toContain('task:task:read');
});

it('exposes the authorization link on the gateway approval card and resumes the user task after consent', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const updates: string[] = [];
  const sdkOptions: unknown[] = [];
  const channel = { rawClient: { request: async (request: import('../office/client').ApiRequest, options: unknown) => { sdkOptions.push(options); requests.push(request); return { code: 0, data: { task: { guid: 't1' } } }; } },
    updateCard: async (_id: string, card: unknown) => { updates.push(JSON.stringify(card)); } } as unknown as import('@larksuiteoapi/node-sdk').LarkChannel;
  const gateway = new FeishuGateway(channel);
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'From bot' });
  await (await store.stateDb()).execute({ sql: 'UPDATE actions SET card_message_id=? WHERE id=?', args: ['auth-card', cards[0]] });
  const event = { operator: { openId: caller.userId }, chatId: caller.chatId, messageId: 'auth-card', action: { value: { approvalId: cards[0], decision: 'approve' } } } as import('@larksuiteoapi/node-sdk').CardActionEvent;
  await gateway.onCardAction(event);
  expect(updates.join('')).toContain('https://accounts.test/authorize');
  const authAction = (await store.getAction(cards[0]))!;
  await permission.completeResourceAuthorization(JSON.parse(authAction.payload).requestId, 'code');
  expect(cards).toHaveLength(2);
  await gateway.onCardAction({ ...event, messageId: 'write-card', action: { tag: 'button', value: { approvalId: cards[1], decision: 'approve' } } });
  expect((await store.getAction(cards[1]))?.status).toBe('succeeded');
  expect(requests).toMatchObject([{ method: 'POST', url: '/open-apis/task/v2/tasks' }]);
  const { withUserAccessToken } = await import('@larksuiteoapi/node-sdk');
  expect(sdkOptions).toEqual([withUserAccessToken('user-token')]);
  await gateway.onCardAction({ ...event, messageId: 'write-card', action: { tag: 'button', value: { approvalId: cards[1], decision: 'approve' } } });
  expect(requests).toHaveLength(1);
});

it('assigns a new personal task to the trusted requester so it is visible in My Tasks', async () => {
  await auth.exchangeCodeForUser('code', 'u1', ['task:task:write']);
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'Mine' });
  const action = JSON.parse((await store.getAction(cards[0]))!.payload);
  await approvals.decideAction(cards[0], caller, true, (_action, key) => permission.executeResourceWrite(caller, action, key));
  expect(requests).toMatchObject([{ data: { members: [{ id: 'u1', type: 'user', role: 'assignee' }] }, userAccessToken: 'user-token' }]);
});

it('rejects an expired OAuth callback before exchanging the authorization code', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_list' });
  const state = await consent();
  await (await store.stateDb()).execute({ sql: 'UPDATE permission_requests SET expires_at=0 WHERE id=?', args: [state] });
  await expect(permission.completeResourceAuthorization(state, 'code')).rejects.toThrow('失效');
  expect(fetch).not.toHaveBeenCalled();
});

it('does not resume when the original message is recalled while code exchange is in flight', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'Cancel in flight' });
  const state = await consent();
  let release!: () => void;
  const hold = new Promise<void>(resolve => { release = resolve; });
  let started!: () => void;
  const entered = new Promise<void>(resolve => { started = resolve; });
  vi.stubGlobal('fetch', vi.fn(async (input: string | URL) => {
    if (String(input).includes('/oauth/token')) { started(); await hold; return new Response(JSON.stringify({ access_token: 'token', expires_in: 3600 })); }
    return new Response(JSON.stringify({ code: 0, data: { open_id: 'u1' } }));
  }));
  const completed = permission.completeResourceAuthorization(state, 'code');
  const rejected = expect(completed).rejects.toThrow('撤回');
  await entered;
  const { FeishuGateway } = await import('../gateway/gateway');
  const gateway = new FeishuGateway({ rawClient: {}, updateCard: async () => undefined } as unknown as import('@larksuiteoapi/node-sdk').LarkChannel);
  await gateway.onMessageRecalled({ messageId: 'origin', chatId: caller.chatId });
  release(); await rejected;
  expect(cards).toHaveLength(1);
  expect(requests).toHaveLength(0);
});

it('does not replay a delete after a network timeout', async () => {
  await auth.exchangeCodeForUser('code', 'u1', ['task:task:write']);
  const { OfficeApi } = await import('../office/client');
  const sent: string[] = [];
  const api = new OfficeApi({ request: async request => { sent.push(request.method); throw new Error('network timeout'); }, send: async () => undefined });
  await expect(permission.executeResourceWrite(caller, { kind: 'task_delete', taskId: 't1', permissionMode: 'user' }, 'key', api)).rejects.toThrow('timeout');
  expect(sent).toEqual(['DELETE']);
  expect(cards).toHaveLength(0);
});

it('detects loss of required scope on token refresh before accessing task data', async () => {
  await auth.exchangeCodeForUser('code', 'u1', ['task:task:read']);
  await (await store.stateDb()).execute("UPDATE user_tokens SET expires_at=0");
  vi.stubGlobal('fetch', vi.fn(async () => new Response(JSON.stringify({ access_token: 'narrow-token', expires_in: 3600, scope: 'calendar:calendar:readonly' }))));
  await permission.requestResourceOperation(caller, { kind: 'task_list' });
  expect(requests).toHaveLength(0);
  expect(cards).toHaveLength(1);
});

it('honors cancellation during resumed write preflight without leaving a writable approval', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_delete', taskId: 't1' });
  const state = await consent();
  const { OfficeApi } = await import('../office/client');
  bridge.setOfficeApi(new OfficeApi({ request: async () => {
    await permission.cancelResourceAuthorization(state);
    return { code: 0, data: { task: { guid: 't1' } } };
  }, send: async () => undefined }));
  await expect(permission.completeResourceAuthorization(state, 'code')).rejects.toThrow('取消');
  const pending = (await (await store.stateDb()).execute("SELECT * FROM actions WHERE kind='task_delete' AND status='pending'")).rows;
  expect(pending).toHaveLength(0);
});

it('reuses the pending write approval when a restarted authorization continuation is retried', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'One only' });
  const state = await consent();
  await permission.completeResourceAuthorization(state, 'code');
  const db = await store.stateDb();
  // Simulate a crash after creating the write card but before completing OAuth state.
  await db.execute({ sql: "UPDATE permission_requests SET status='resource_auth_processing' WHERE id=?", args: [state] });
  await store.recoverInterruptedApprovals();
  await permission.resolveResourcePermissionReply(caller, '重新授权');
  const next = (await store.getAction(cards.at(-1)!))!;
  const authAction = JSON.parse(next.payload);
  await approvals.decideAction(next.id, caller, true, (_action, key) => permission.beginResourceAuthorization(authAction, key));
  await permission.completeResourceAuthorization(authAction.requestId, 'code2');
  const pending = (await db.execute("SELECT * FROM actions WHERE kind='task_create' AND status='pending'")).rows;
  expect(pending).toHaveLength(1);
});

it('routes calendar and task read tools through authorization and registers all CRUD tools', async () => {
  const { officeTools } = await import('../mastra/tools/office-tools');
  for (const name of ['listCalendarsTool', 'readCalendarInfoTool', 'readCalendarEventTool', 'deleteCalendarTool', 'createCalendarContainerTool', 'updateCalendarContainerTool', 'deleteCalendarContainerTool', 'listTasksTool', 'deleteTaskTool']) expect(officeTools).toHaveProperty(name);
  const context = { requestContext: { get: (key: string) => key === 'caller' ? caller : undefined } };
  await officeTools.readTaskTool.execute!({ taskId: 't1' }, context as never);
  expect(cards).toHaveLength(1);
  expect(requests).toHaveLength(0);
});

it('keeps the original write card waiting for authorization on duplicate clicks after a token expires', async () => {
  await auth.exchangeCodeForUser('code', 'u1', ['task:task:write']);
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'Token expires before confirm' });
  const actionId = cards[0];
  await (await store.stateDb()).execute('DELETE FROM user_tokens');
  const { FeishuGateway } = await import('../gateway/gateway');
  const updates: string[] = [];
  const gateway = new FeishuGateway({ rawClient: {}, updateCard: async (_id: string, card: unknown) => { updates.push(JSON.stringify(card)); } } as unknown as import('@larksuiteoapi/node-sdk').LarkChannel);
  const event = { operator: { openId: caller.userId }, chatId: caller.chatId, messageId: 'write-card', action: { tag: 'button', value: { approvalId: actionId, decision: 'approve' } } } as import('@larksuiteoapi/node-sdk').CardActionEvent;
  await gateway.onCardAction(event);
  await gateway.onCardAction(event);
  expect(updates).toHaveLength(2);
  expect(updates.every(card => card.includes('等待用户授权，原操作尚未执行'))).toBe(true);
  expect(cards).toHaveLength(2);
  expect(requests).toHaveLength(0);
});

it('corrects the reported task date and preserves both endpoints through authorization and execution', async () => {
  const { createTaskTool } = await import('../mastra/tools/office-tools');
  const { officeAction, actionSummary } = await import('../office/actions');
  const context = { requestContext: { get: (key: string) => ({ caller,
    originalMessage: '创建一个任务，今天下午16:00-17:00，主题是完成重复规则需求不明确的规则的梳理和编写',
    messageCreatedAt: Date.parse('2026-11-17T15:16:01.581+08:00'),
  })[key] } };
  const execute = Reflect.get(createTaskTool, 'execute') as (input: unknown, context: unknown) => Promise<{ requestId: string }>;
  const result = await execute({ summary: '完成重复规则需求不明确的规则的梳理和编写',
    description: '时间窗口：2026-10-26 16:00–17:00（Asia/Shanghai）', dueTimestamp: '1793005200000', permissionMode: 'user' }, context);
  const row = (await (await store.stateDb()).execute({ sql: 'SELECT payload FROM permission_requests WHERE id=?', args: [result.requestId] })).rows[0];
  const expected = { startTimestamp: String(Date.parse('2026-11-17T16:00:00+08:00')), dueTimestamp: String(Date.parse('2026-11-17T17:00:00+08:00')) };
  expect(JSON.parse(String(row.payload)).intent).toMatchObject(expected);
  expect(JSON.parse(String(row.payload)).intent.description).not.toContain('2026-10-26');
  await permission.completeResourceAuthorization(await consent(), 'code');
  const write = officeAction.parse(JSON.parse((await store.getAction(cards[1]))!.payload));
  expect(write).toMatchObject(expected);
  expect(actionSummary(write)).toContain('2026-11-17 16:00');
  expect(actionSummary(write)).toContain('2026-11-17 17:00');
  await approvals.decideAction(cards[1], caller, true, (action, key) => {
    if (action.kind !== 'task_create') throw new Error('expected task');
    return permission.executeResourceWrite(caller, action, key);
  });
  expect(requests).toMatchObject([{ method: 'POST', data: {
    start: { timestamp: expected.startTimestamp, is_all_day: false }, due: { timestamp: expected.dueTimestamp, is_all_day: false },
  } }]);
});

it('anchors today to the original Shanghai message day even when OAuth finishes after midnight', async () => {
  const receivedAt = Date.parse('2026-11-17T23:55:00+08:00');
  const now = vi.spyOn(Date, 'now').mockReturnValue(receivedAt);
  try {
    await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'Original day', dueTimestamp: '1793005200000' },
      { originalMessage: '创建今天16:00-17:00的任务', messageCreatedAt: receivedAt });
    const state = await consent();
    now.mockReturnValue(receivedAt + 10 * 60_000);
    await permission.completeResourceAuthorization(state, 'code');
    expect(JSON.parse((await store.getAction(cards[1]))!.payload)).toMatchObject({
      startTimestamp: String(Date.parse('2026-11-17T16:00:00+08:00')),
      dueTimestamp: String(Date.parse('2026-11-17T17:00:00+08:00')),
    });
  } finally { now.mockRestore(); }
});

it('normalizes task times when a user already has authorization', async () => {
  await auth.exchangeCodeForUser('code', 'u1', ['task:task:write']);
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'Already authorized', dueTimestamp: '1793005200000' },
    { originalMessage: '创建今天16:00-17:00的任务', messageCreatedAt: Date.parse('2026-11-17T15:16:00+08:00') });
  expect(JSON.parse((await store.getAction(cards[0]))!.payload)).toMatchObject({
    kind: 'task_create', startTimestamp: String(Date.parse('2026-11-17T16:00:00+08:00')), dueTimestamp: String(Date.parse('2026-11-17T17:00:00+08:00')),
  });
});

it('repairs a legacy pending intent from its saved source before presenting the write approval', async () => {
  await permission.requestResourceOperation(caller, { kind: 'task_create', summary: 'Saved before fix', dueTimestamp: '1793005200000' });
  const state = await consent();
  const db = await store.stateDb();
  const row = (await db.execute({ sql: 'SELECT payload FROM permission_requests WHERE id=?', args: [state] })).rows[0];
  const saved = JSON.parse(String(row.payload));
  saved.continuation = { originalMessage: '创建今天16:00-17:00的任务', messageCreatedAt: Date.parse('2026-11-17T15:16:00+08:00') };
  await db.execute({ sql: 'UPDATE permission_requests SET payload=? WHERE id=?', args: [JSON.stringify(saved), state] });
  await permission.completeResourceAuthorization(state, 'code');
  expect(JSON.parse((await store.getAction(cards[1]))!.payload)).toMatchObject({
    startTimestamp: String(Date.parse('2026-11-17T16:00:00+08:00')), dueTimestamp: String(Date.parse('2026-11-17T17:00:00+08:00')),
  });
});

it('resumes colleague search after OAuth and approves attendee creation exactly once', async () => {
  const resumed: unknown[] = [];
  permission.setResourceContinuation(async (...args) => { resumed.push(args); });
  const source = { originalMessage: '把张三加入刚创建的日程', messageCreatedAt: Date.now() };
  await permission.requestResourceOperation(caller, { kind: 'openapi_read', operationId: 'contact.user.search', permissionMode: 'user', scopes: [], summary: '搜索张三', arguments: { params: { query: '张三' } } }, source);
  expect(requests).toHaveLength(0);
  await permission.completeResourceAuthorization(await consent(), 'code');
  expect(requests[0]).toMatchObject({ method: 'GET', url: '/open-apis/search/v1/user', params: { query: '张三' }, userAccessToken: 'user-token' });
  expect(resumed[0]).toMatchObject([caller, source, expect.any(Object), expect.any(String)]);
  await auth.exchangeCodeForUser('code', 'u1', ['contact:user:search', 'calendar:calendar']);
  await permission.requestResourceOperation(caller, { kind: 'openapi_write', operationId: 'calendar.v4.calendar.event.attendee.create', permissionMode: 'user', scopes: [], summary: '邀请选定同事张三', arguments: {
    path: { calendar_id: 'cal', event_id: 'event1' }, params: { user_id_type: 'open_id' }, data: { attendees: [{ type: 'user', user_id: 'ou_zhangsan' }] },
  } }, source);
  expect(requests).toHaveLength(1);
  const id = cards.at(-1)!;
  const execute = (action: import('../office/actions').OfficeAction, key: string) => {
    if (action.kind !== 'openapi_write') throw new Error('Expected API write');
    return permission.executeResourceWrite(caller, action, key);
  };
  await approvals.decideAction(id, caller, true, execute);
  await approvals.decideAction(id, caller, true, execute);
  expect(requests).toHaveLength(2);
  expect(requests[1]).toMatchObject({ method: 'POST', url: '/open-apis/calendar/v4/calendars/cal/events/event1/attendees', userAccessToken: 'user-token', data: { attendees: [{ type: 'user', user_id: 'ou_zhangsan' }] } });
});

it('uses application grants for directory reads and does not replace an app denial with user OAuth', async () => {
  const input = { kind: 'openapi_read' as const, operationId: 'contact.v3.user.get', permissionMode: 'app' as const, scopes: [], summary: '查询用户', arguments: { path: { user_id: 'ou1' }, params: { user_id_type: 'open_id' } } };
  await permission.requestResourceOperation(caller, input);
  expect(requests[0]).toMatchObject({ method: 'GET', url: '/open-apis/contact/v3/users/ou1', userAccessToken: undefined });
  expect(cards).toHaveLength(0);
  responseCode = 99991672;
  await expect(permission.requestResourceOperation(caller, input)).rejects.toThrow('管理员');
  expect(cards).toHaveLength(0);
});

it('rejects a generic write disguised as a read before authorization or API access', async () => {
  await expect(permission.requestResourceOperation(caller, { kind: 'openapi_read', operationId: 'calendar.v4.calendar.event.attendee.create', permissionMode: 'user', scopes: [], summary: 'Read?', arguments: { path: { calendar_id: 'cal', event_id: 'evt' } } })).rejects.toThrow('读写类型');
  expect(cards).toHaveLength(0);
  expect(requests).toHaveLength(0);
});

it('corrects calendar clocks before OAuth and preserves them through the write approval', async () => {
  const ms = (clock: string) => Date.parse(`2026-11-18T${clock}:00+08:00`);
  await permission.requestResourceOperation(caller, { kind: 'calendar_create', calendarId: 'cal', summary: '时间核对', startTime: String(ms('21:40') / 1000), endTime: String(ms('22:40') / 1000) }, { originalMessage: '创建今天11:00-12:00的日程', messageCreatedAt: ms('10:00') });
  await permission.completeResourceAuthorization(await consent(), 'code');
  const write = JSON.parse((await store.getAction(cards[1]))!.payload);
  expect(write).toMatchObject({ startTime: String(ms('11:00') / 1000), endTime: String(ms('12:00') / 1000) });
  await approvals.decideAction(cards[1], caller, true, (action, key) => permission.executeResourceWrite(caller, action as never, key));
  expect(requests.at(-1)).toMatchObject({ method: 'POST', data: { start_time: { timestamp: String(ms('11:00') / 1000) }, end_time: { timestamp: String(ms('12:00') / 1000) } } });
});

it('returns current capabilities and isolates recently created events by requester and chat', async () => {
  const { officeTools } = await import('../mastra/tools/office-tools');
  const context = { requestContext: { get: (key: string) => key === 'caller' ? caller : undefined } };
  const invoke = (name: keyof typeof officeTools) => (Reflect.get(officeTools[name], 'execute') as (input: unknown, context: unknown) => Promise<Record<string, unknown>>)({}, context);
  const capabilities = JSON.stringify(await invoke('capabilitiesTool'));
  expect(capabilities).toContain('updateCalendarTool');
  expect(capabilities).toContain('群聊消息');
  expect(capabilities).toContain('office_call_api');
  expect(capabilities).toContain('start_schedule');
  await auth.exchangeCodeForUser('code', 'u1', ['calendar:calendar']);
  await permission.requestResourceOperation(caller, { kind: 'calendar_create', calendarId: 'cal', summary: 'My event', startTime: '1794970800', endTime: '1794974400' });
  await approvals.decideAction(cards[0], caller, true, async () => ({ event: { event_id: 'evt_real' } }));
  expect(await invoke('recentCalendarEventsTool')).toMatchObject({ events: [{ calendarId: 'cal', eventId: 'evt_real' }] });
  await (await store.stateDb()).execute({ sql: 'UPDATE actions SET requester_id=? WHERE id=?', args: ['someone_else', cards[0]] });
  expect(await invoke('recentCalendarEventsTool')).toEqual({ events: [] });
});

it('uses actual missing scopes to request incremental authorization without repeating an identical grant', async () => {
  const { OfficeApi } = await import('../office/client');
  bridge.setOfficeApi(new OfficeApi({ request: async () => ({ code: 99991672, msg: 'Required scope: contact:user.base:readonly' }), send: async () => undefined }));
  await auth.exchangeCodeForUser('code', 'u1', ['offline_access']);
  const intent = { kind: 'openapi_read' as const, operationId: 'contact.v3.user.get', permissionMode: 'user' as const, scopes: [], summary: '查询用户', arguments: { path: { user_id: 'ou1' } } };
  const outcome = await permission.requestResourceOperation(caller, intent);
  expect(outcome.status).toBe('authorization_required');
  const row = (await (await store.stateDb()).execute('SELECT payload FROM permission_requests')).rows[0];
  expect(JSON.parse(String(row.payload)).intent.scopes).toEqual(['contact:user.base:readonly']);
  await auth.exchangeCodeForUser('code', 'u1', ['contact:user.base:readonly']);
  await expect(permission.requestResourceOperation(caller, { ...intent, scopes: ['contact:user.base:readonly'] })).rejects.toThrow('不会重复');
  expect(cards).toHaveLength(1);
});

it('continues a read OAuth flow with a newly discovered scope while preserving the original request', async () => {
  const { OfficeApi } = await import('../office/client');
  bridge.setOfficeApi(new OfficeApi({ request: async () => ({ code: 99991672, msg: 'Required scope: contact:user.base:readonly' }), send: async () => undefined }));
  const source = { originalMessage: '查询同事', messageCreatedAt: Date.now() };
  await permission.requestResourceOperation(caller, { kind: 'openapi_read', operationId: 'contact.v3.user.get', permissionMode: 'user', scopes: [], summary: '查询用户', arguments: { path: { user_id: 'ou1' } } }, source);
  await permission.completeResourceAuthorization(await consent(), 'code');
  const rows = (await (await store.stateDb()).execute("SELECT payload FROM permission_requests WHERE status='resource_approval'")).rows;
  expect(rows).toHaveLength(1);
  expect(JSON.parse(String(rows[0].payload))).toMatchObject({ intent: { scopes: ['contact:user.base:readonly'] }, continuation: source });
});

async function unified(operationId: string, args: Record<string, unknown>) {
  const { callFeishuApiTool } = await import('../mastra/tools/office-tools');
  const execute = Reflect.get(callFeishuApiTool, 'execute') as (input: unknown, context: unknown) => Promise<unknown>;
  return execute({ operationId, arguments: args }, { requestContext: { get: (key: string) => key === 'caller' ? caller : undefined } });
}

it('corrects the reported January date at the unified boundary before approval and the actual calendar write', async () => {
  const { callFeishuApiTool } = await import('../mastra/tools/office-tools');
  const { RequestContext } = await import('@mastra/core/request-context');
  await auth.exchangeCodeForUser('code', 'u1', ['calendar:calendar']);
  const requestContext = new RequestContext();
  requestContext.set('caller', caller);
  requestContext.set('originalMessage', '创建一个日程，主题是xxx，时间为今天18:00-19:00');
  requestContext.set('messageCreatedAt', Date.parse('2026-11-20T10:00:00+08:00'));
  const execute = Reflect.get(callFeishuApiTool, 'execute') as (input: unknown, context: unknown) => Promise<unknown>;
  await execute({ operationId: 'office_calendar_create', arguments: { calendarId: 'cal', summary: 'xxx', permissionMode: 'user',
    startTime: String(Date.parse('2027-01-15T11:20:00+08:00') / 1000), endTime: String(Date.parse('2027-01-15T12:20:00+08:00') / 1000) } }, { requestContext });
  expect(cards).toHaveLength(1);
  expect(JSON.parse((await store.getAction(cards[0]))!.payload)).toMatchObject({ startTime: '1795168800', endTime: '1795172400' });
  expect(requests.every(request => request.method === 'GET')).toBe(true);
  await approvals.decideAction(cards[0], caller, true, (action, key) => permission.executeResourceWrite(caller, action as never, key));
  expect(requests.at(-1)).toMatchObject({ method: 'POST', data: { start_time: { timestamp: '1795168800' }, end_time: { timestamp: '1795172400' } } });
});

it.each([
  ['office_calendar_create', { calendarId: 'cal', summary: 'New', startTime: '1794970800', endTime: '1794974400', permissionMode: 'app' }, 'calendar_create'],
  ['office_calendar_update', { calendarId: 'cal', eventId: 'evt', summary: 'Updated', permissionMode: 'app' }, 'calendar_update'],
  ['office_task_create', { summary: 'New', permissionMode: 'app' }, 'task_create'],
  ['office_task_update', { taskId: 'task1', summary: 'Updated', permissionMode: 'app' }, 'task_update'],
  ['office_doc_create', { title: 'Doc', content: 'Text' }, 'doc_create'],
  ['office_doc_update', { documentId: 'doxcn123', blockId: 'block', text: 'Text' }, 'doc_update'],
  ['office_drive_create_folder', { parentFolderUrl: 'https://api.test/drive/folder/folder1', name: 'Child' }, 'drive_folder_create'],
  ['office_drive_copy_file', { fileToken: 'doxcn123', fileType: 'docx', name: 'Copy', destinationFolderUrl: 'https://api.test/drive/folder/folder1' }, 'drive_file_copy'],
  ['office_drive_file_rename', { fileToken: 'doxcn123', fileType: 'docx', folderToken: 'folder1', expectedName: 'Old', newName: 'New' }, 'drive_file_rename'],
  ['office_message_send', { targetChatId: 'dm1', text: 'Hello' }, 'message_send'],
])('routes existing business operation %s through its original approval implementation', async (operationId, args, kind) => {
  const result = await unified(operationId as string, args as Record<string, unknown>);
  expect(result).toMatchObject({ status: 'pending' });
  expect(cards).toHaveLength(1);
  const row = await store.getAction(cards[0]);
  expect(row?.kind).toBe(kind);
  expect(requests.every(request => request.method === 'GET')).toBe(true);
});

it('exposes every business operation schema through discovery and removes duplicate model tools', async () => {
  const { officeBusinessTools, discoverOfficeOperations } = await import('../mastra/tools/office-tools');
  for (const tool of Object.values(officeBusinessTools)) {
    expect(discoverOfficeOperations('', tool.id)).toMatchObject({ operation: { id: tool.id, input: { type: 'object' }, kind: 'business' } });
  }
  const { agent } = await import('../mastra/agents/agent');
  const tools = await agent.listTools();
  expect(tools).toHaveProperty('office_call_api');
  expect(tools).toHaveProperty('office_discover_apis');
  for (const name of Object.keys(officeBusinessTools)) expect(tools).not.toHaveProperty(name);
  expect(JSON.stringify(discoverOfficeOperations('表格'))).toContain('office_table_');
  expect(discoverOfficeOperations('创建任务').operations).toContainEqual(expect.objectContaining({ id: 'office_task_create', kind: 'business' }));
});

it('preserves table preflight through the unified entry and checks changed cells before writing', async () => {
  const { OfficeApi } = await import('../office/client');
  let value = 'Old';
  bridge.setOfficeApi(new OfficeApi({ request: async request => {
    requests.push(request);
    if (request.url.endsWith('/members/auth')) return { code: 0, data: { auth_result: true } };
    if (request.url.endsWith('/sheets/query')) return { code: 0, data: { sheets: [{ sheet_id: 'sheet1', title: 'Sheet' }] } };
    if (request.url.includes('/values/')) return { code: 0, data: { valueRange: { values: [[value]] } } };
    return { code: 0, data: {} };
  }, send: async () => undefined }));
  expect(await unified('office_table_inspect', { url: 'https://api.test/sheets/token1', permissionMode: 'app' })).toMatchObject({ status: 'read' });
  const result = await unified('office_table_update_cell', { url: 'https://api.test/sheets/token1', cell: 'A2', value: 'New', permissionMode: 'app' });
  expect(result).toMatchObject({ status: 'pending' });
  const row = await store.getAction(cards[0]);
  expect(JSON.parse(row!.payload)).toMatchObject({ kind: 'sheet_cell_set', expectedValue: 'Old', value: 'New' });
  value = 'Changed after approval';
  const outcome = await approvals.decideAction(cards[0], caller, true, (action, key) => bridge.getOfficeApi().execute(action, key));
  expect(outcome.status).toBe('failed');
  expect(requests.every(request => request.method === 'GET')).toBe(true);
});

it('preserves the target-chat allowlist when a message is sent through the unified entry', async () => {
  const { OfficeApi } = await import('../office/client');
  const send = vi.fn();
  bridge.setOfficeApi(new OfficeApi({ request: async () => ({ code: 0, data: {} }), send }));
  await unified('office_message_send', { targetChatId: 'not-allowlisted', text: 'Hello' });
  const outcome = await approvals.decideAction(cards[0], caller, true, (action, key) => bridge.getOfficeApi().execute(action, key));
  expect(outcome.status).toBe('failed');
  expect(send).not.toHaveBeenCalled();
  await expect(unified('im.v1.message.create', { params: { receive_id_type: 'chat_id' }, data: { receive_id: 'not-allowlisted', msg_type: 'text', content: '{}' } })).rejects.toThrow();
});

it('reads documents, tasks, calendar events and folders through the same business dispatcher', async () => {
  const { OfficeApi } = await import('../office/client');
  bridge.setOfficeApi(new OfficeApi({ request: async request => { requests.push(request); return { code: 0, data: { files: [], items: [], has_more: false, calendar: { calendar_id: 'cal', role: 'owner' } } }; }, send: async () => undefined }));
  await auth.exchangeCodeForUser('code', 'u1', ['calendar:calendar:readonly', 'task:task:read']);
  await unified('office_read_document', { documentId: 'doxcn123' });
  await unified('office_read_task', { taskId: 'task1' });
  await unified('office_read_calendar', { calendarId: 'cal', startTime: '1794970800', endTime: '1794974400' });
  await unified('office_drive_list_folder', { folderUrl: 'https://api.test/drive/folder/folder1' });
  expect(requests.map(request => request.url)).toEqual(['/open-apis/docx/v1/documents/doxcn123/raw_content', '/open-apis/task/v2/tasks/task1', '/open-apis/calendar/v4/calendars/cal/events', '/open-apis/drive/v1/files']);
  expect(cards).toHaveLength(0);
});

it('preserves Bitable field preflight through the unified entry', async () => {
  const { OfficeApi } = await import('../office/client');
  bridge.setOfficeApi(new OfficeApi({ request: async request => {
    requests.push(request);
    if (request.url.endsWith('/members/auth')) return { code: 0, data: { auth_result: true } };
    if (request.url.endsWith('/tables')) return { code: 0, data: { items: [{ table_id: 'tbl1', name: 'Table' }] } };
    if (request.url.endsWith('/fields')) return { code: 0, data: { items: [{ field_name: 'Title', type: 1 }] } };
    return { code: 0, data: {} };
  }, send: async () => undefined }));
  await expect(unified('office_table_create_record', { url: 'https://api.test/base/base1', fields: { Unknown: 'x' }, permissionMode: 'app' })).rejects.toThrow();
  expect(cards).toHaveLength(0);
  expect(await unified('office_table_create_record', { url: 'https://api.test/base/base1', fields: { Title: 'x' }, permissionMode: 'app' })).toMatchObject({ status: 'pending' });
  expect((await store.getAction(cards[0]))?.kind).toBe('bitable_record_create');
  expect(requests.every(request => request.method === 'GET')).toBe(true);
});
