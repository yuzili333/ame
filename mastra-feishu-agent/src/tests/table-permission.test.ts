import { beforeAll, afterAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { LarkChannel, NormalizedMessage } from '@larksuiteoapi/node-sdk';

vi.mock('../mastra', () => ({ controller: {} }));

process.env.AGENT_DATA_DIR = mkdtempSync(join(tmpdir(), 'mastra-feishu-table-test-'));
process.env.FEISHU_DOCUMENT_HOSTS = 'feishu.cn';
process.env.FEISHU_APP_ID = 'cli_test';
process.env.FEISHU_APP_SECRET = 'test-secret';
process.env.FEISHU_DOMAIN = 'https://tenant.feishu.cn';
process.env.FEISHU_OAUTH_AUTHORIZE_URL = 'https://accounts.feishu.cn/oauth/authorize';
process.env.FEISHU_OAUTH_TOKEN_URL = 'https://accounts.feishu.cn/oauth/token';
process.env.FEISHU_OAUTH_REDIRECT_URI = 'http://localhost:4000/auth/feishu/callback';
process.env.FEISHU_TOKEN_ENCRYPTION_KEY = 'test-encryption-key';

let office: typeof import('../office/client');
let bridge: typeof import('../runtime/office-bridge');
let approvalBridge: typeof import('../runtime/approval-bridge');
let tables: typeof import('../office/table-permission');
let approvals: typeof import('../state/approvals');
let store: typeof import('../state/store');
let auth: typeof import('../office/user-auth');

beforeAll(async () => {
  office = await import('../office/client');
  bridge = await import('../runtime/office-bridge');
  approvalBridge = await import('../runtime/approval-bridge');
  tables = await import('../office/table-permission');
  approvals = await import('../state/approvals');
  store = await import('../state/store');
  auth = await import('../office/user-auth');
});
afterAll(() => {
  bridge.setOfficeApi(undefined);
  approvalBridge.setApprovalNotifier(undefined);
  tables.setTablePermissionNotifier(undefined);
  vi.restoreAllMocks();
});

const caller = { userId: 'u-table', chatId: 'g-table', chatType: 'group' as const };
const intent = { operation: 'fill_first_empty' as const, url: 'https://tenant.feishu.cn/sheets/sht1?sheet=s1', columnHeader: '名称', value: 'Topson' };
function approvalId(result: Awaited<ReturnType<typeof tables.requestTableOperation>>): string {
  if (!('approvalId' in result)) throw new Error('Expected a pending approval');
  return result.approvalId;
}

function writableApi() {
  let writes = 0;
  const api = new office.OfficeApi({
    request: async request => {
      if (request.url.includes('/members/auth')) return { code: 0, data: { auth_result: true } };
      if (request.url.includes('/sheets/query')) return { code: 0, data: { sheets: [{ sheet_id: 's1', title: 'Main', grid_properties: { row_count: 1000 } }] } };
      if (request.url.includes('/values/')) {
        const range = decodeURIComponent(request.url.split('/values/')[1]);
        return { code: 0, data: { valueRange: { values: range === 's1!A1:IV1' ? [['名称']] : [] } } };
      }
      if (request.method === 'PUT') writes++;
      return { code: 0, data: {} };
    },
    send: async () => undefined,
  });
  return { api, count: () => writes };
}

describe('table permission and persistent approval', () => {
  it('requests table scopes and distinguishes calendar-only from table-authorized user tokens', async () => {
    const scopes = ['calendar:calendar', 'sheets:spreadsheet:read', 'sheets:spreadsheet:write_only', 'offline_access'];
    expect(new URL(auth.authorizationUrl('state1', scopes)).searchParams.get('scope')).toBe(scopes.join(' '));
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockImplementation(async input => {
      if (String(input).includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'u-token', refresh_token: 'r-token', expires_in: 3600, scope: 'calendar:calendar sheets:spreadsheet:read sheets:spreadsheet:write_only' }), { status: 200 });
      return new Response(JSON.stringify({ code: 0, data: { open_id: 'u-scope' } }), { status: 200 });
    });
    try {
      await auth.exchangeCodeForUser('code1', 'u-scope', scopes);
      expect(await auth.userAccessToken('u-scope', ['sheets:spreadsheet:read', 'sheets:spreadsheet:write_only'])).toBe('u-token');
      expect(await auth.userAccessToken('u-scope', ['bitable:app'])).toBeNull();
    } finally { fetchMock.mockRestore(); }
  });

  it('keeps a group requester attached to a permission request and permits cancellation without a card', async () => {
    bridge.setOfficeApi(new office.OfficeApi({ request: async () => ({ code: 99991672, msg: 'missing wiki:node:read' }), send: async () => undefined }));
    const notices: string[] = [];
    tables.setTablePermissionNotifier(async (_chat, text) => { notices.push(text); });
    const result = await tables.requestTableOperation(caller, { ...intent, url: 'https://tenant.feishu.cn/wiki/node1' });
    expect(result.status).toBe('permission_choice_required');
    if (!('requestId' in result)) throw new Error('Expected a permission request');
    const row = (await (await store.stateDb()).execute({ sql: 'SELECT chat_type, status FROM permission_requests WHERE id = ?', args: [result.requestId] })).rows[0];
    expect(row).toMatchObject({ chat_type: 'group', status: 'table_choice' });
    expect(await tables.resolveTablePermissionReply(caller, '取消')).toBe(true);
    expect(notices.at(-1)).toContain('已取消');
  });

  it('asks for user permission before any approval when the app can read but cannot edit', async () => {
    bridge.setOfficeApi(new office.OfficeApi({
      request: async request => request.url.includes('/members/auth')
        ? { code: 0, data: { auth_result: false } }
        : { code: 0, data: {} },
      send: async () => undefined,
    }));
    const notices: string[] = [];
    const cards: string[] = [];
    tables.setTablePermissionNotifier(async (_chat, text) => { notices.push(text); });
    approvalBridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
    const result = await tables.requestTableOperation(caller, intent);
    expect(result.status).toBe('permission_choice_required');
    expect(notices[0]).toContain('没有编辑权限');
    expect(cards).toHaveLength(0);
    await tables.resolveTablePermissionReply(caller, '取消');
  });

  it('retries app permission after a contextual B reply and prepares one approval', async () => {
    bridge.setOfficeApi(new office.OfficeApi({ request: async () => ({ code: 99991672, msg: 'missing sheets scope' }), send: async () => undefined }));
    tables.setTablePermissionNotifier(async () => undefined);
    const pending = await tables.requestTableOperation(caller, { ...intent, value: 'B-approved' });
    expect(pending.status).toBe('permission_choice_required');
    const { api } = writableApi();
    bridge.setOfficeApi(api);
    const cards: string[] = [];
    approvalBridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
    expect(await tables.resolveTablePermissionReply(caller, 'B，已开通应用操作云文档权限，请重试')).toBe(true);
    expect(cards).toHaveLength(1);
    expect(await tables.resolveTablePermissionReply(caller, '应用权限')).toBe(false);
  });

  it('supersedes a stale permission choice when a complete table request is resent', async () => {
    bridge.setOfficeApi(new office.OfficeApi({ request: async () => ({ code: 99991672, msg: 'missing sheets scope' }), send: async () => undefined }));
    tables.setTablePermissionNotifier(async () => undefined);
    const pending = await tables.requestTableOperation(caller, intent);
    if (!('requestId' in pending)) throw new Error('Expected a permission request');
    const message = '在表格表头列“名称”下方第一个空白单元格写入“Topson”，表格链接为[表格测试文件](https://tenant.feishu.cn/wiki/node1)';
    expect(await tables.supersedeTablePermissionForNewRequest(caller, message)).toBe(true);
    const row = (await (await store.stateDb()).execute({ sql: 'SELECT status FROM permission_requests WHERE id = ?', args: [pending.requestId] })).rows[0];
    expect(row?.status).toBe('superseded');
    expect(await tables.resolveTablePermissionReply(caller, message)).toBe(false);
    expect(await tables.supersedeTablePermissionForNewRequest(caller, message)).toBe(true);
    expect(await tables.supersedeTablePermissionForNewRequest(caller, '应用权限')).toBe(false);
  });

  it('routes an explicit first-empty-cell request directly to one approval', async () => {
    const { parseFirstEmptySheetCellRequest } = await import('../office/table-request');
    const { FeishuGateway } = await import('../gateway/gateway');
    const content = '在表格表头列“名称”下方第一个空白单元格写入“Direct test”，表格链接为[测试](https://tenant.feishu.cn/sheets/sht1?sheet=s1)';
    expect(parseFirstEmptySheetCellRequest(content)).toMatchObject({ operation: 'fill_first_empty', columnHeader: '名称', value: 'Direct test', permissionMode: 'app' });
    expect(parseFirstEmptySheetCellRequest('请修改表格')).toBeUndefined();
    const { api } = writableApi();
    bridge.setOfficeApi(api);
    const cards: string[] = [];
    approvalBridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
    const gateway = new FeishuGateway({ send: async () => ({}), rawClient: { request: async () => ({ code: 0, data: {} }) } } as unknown as LarkChannel);
    const message = { messageId: crypto.randomUUID(), senderId: caller.userId, chatId: caller.chatId, chatType: caller.chatType,
      content, mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: Date.now() } as NormalizedMessage;
    await gateway.onMessage(message);
    expect(cards).toHaveLength(1);
    expect((await store.getAction(cards[0]))?.kind).toBe('sheet_cell_set');
    await gateway.onMessage(message);
    expect(cards).toHaveLength(1);
  });

  it('creates only one card for a repeated proposal and executes one confirmed write', async () => {
    const { api, count } = writableApi();
    bridge.setOfficeApi(api);
    const cards: string[] = [];
    approvalBridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
    const first = await tables.requestTableOperation(caller, intent);
    const again = await tables.requestTableOperation(caller, intent);
    expect(first).toMatchObject({ status: 'pending' });
    expect(again).toMatchObject({ approvalId: approvalId(first) });
    expect(cards).toHaveLength(1);
    expect((await approvals.decideAction(approvalId(first), { userId: 'another', chatId: caller.chatId }, true, action => api.execute(action))).status).toBe('denied');
    expect((await approvals.decideAction(approvalId(first), caller, true, action => api.execute(action))).status).toBe('succeeded');
    expect((await approvals.decideAction(approvalId(first), caller, true, action => api.execute(action))).status).toBe('succeeded');
    expect(count()).toBe(1);
  });

  it('does not write after a denied or expired table approval', async () => {
    const { api, count } = writableApi();
    bridge.setOfficeApi(api);
    approvalBridge.setApprovalNotifier(async () => undefined);
    const denied = await tables.requestTableOperation(caller, { ...intent, value: 'Denied' });
    expect((await approvals.decideAction(approvalId(denied), caller, false, action => api.execute(action))).status).toBe('denied');
    const expired = await tables.requestTableOperation(caller, { ...intent, value: 'Expired' });
    await (await store.stateDb()).execute({ sql: 'UPDATE actions SET expires_at = 0 WHERE id = ?', args: [approvalId(expired)] });
    expect((await approvals.decideAction(approvalId(expired), caller, true, action => api.execute(action))).status).toBe('expired');
    expect(count()).toBe(0);
  });

  it('does not replay a table write after an uncertain API failure', async () => {
    let writes = 0;
    const api = new office.OfficeApi({
      request: async request => {
        if (request.url.includes('/members/auth')) return { code: 0, data: { auth_result: true } };
        if (request.url.includes('/sheets/query')) return { code: 0, data: { sheets: [{ sheet_id: 's1', title: 'Main' }] } };
        if (request.url.includes('/values/')) return { code: 0, data: { valueRange: { values: decodeURIComponent(request.url).includes('A1:IV1') ? [['名称']] : [] } } };
        if (request.method === 'PUT') { writes++; throw new Error('upstream timeout after commit'); }
        return { code: 0, data: {} };
      },
      send: async () => undefined,
    });
    bridge.setOfficeApi(api);
    approvalBridge.setApprovalNotifier(async () => undefined);
    const proposed = await tables.requestTableOperation(caller, { ...intent, value: 'uncertain' });
    const id = approvalId(proposed);
    expect((await approvals.decideAction(id, caller, true, action => api.execute(action))).status).toBe('failed');
    expect((await approvals.decideAction(id, caller, true, action => api.execute(action))).status).toBe('failed');
    expect(writes).toBe(1);
  });

  it('claims an OAuth callback once and keeps the original group context for approval', async () => {
    const { api } = writableApi();
    bridge.setOfficeApi(api);
    const cards: string[] = [];
    approvalBridge.setApprovalNotifier(async (_chat, id) => { cards.push(id); });
    const state = crypto.randomUUID();
    await (await store.stateDb()).execute({ sql: `INSERT INTO permission_requests(id, requester_id, chat_id, chat_type, payload, status, created_at, expires_at) VALUES (?, ?, ?, ?, ?, 'table_auth', ?, ?)`, args: [state, caller.userId, caller.chatId, caller.chatType, JSON.stringify(intent), Date.now(), Date.now() + 60_000] });
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockImplementation(async input => {
      if (String(input).includes('/oauth/token')) return new Response(JSON.stringify({ code: 0, access_token: 'u-table-token', refresh_token: 'r-table-token', expires_in: 3600 }), { status: 200 });
      return new Response(JSON.stringify({ code: 0, data: { open_id: caller.userId } }), { status: 200 });
    });
    try {
      await tables.completeTableAuthorization(state, 'code1');
      await expect(tables.completeTableAuthorization(state, 'code1')).rejects.toThrow('已失效或已处理');
      expect(cards).toHaveLength(1);
      expect((await store.getAction(cards[0]))?.chat_type).toBe('group');
    } finally { fetchMock.mockRestore(); }
  });
});
