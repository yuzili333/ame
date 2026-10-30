import { beforeAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { CardActionEvent, LarkChannel, NormalizedMessage } from '@larksuiteoapi/node-sdk';

process.env.AGENT_DATA_DIR = mkdtempSync(join(tmpdir(), 'mastra-feishu-test-'));
process.env.FEISHU_ALLOWED_USERS = 'u1,u2';
process.env.FEISHU_ALLOWED_CHATS = 'dm1,g1,target';
process.env.FEISHU_OAUTH_AUTHORIZE_URL = '';
process.env.FEISHU_OAUTH_TOKEN_URL = '';
process.env.FEISHU_OAUTH_REDIRECT_URI = '';
process.env.FEISHU_TOKEN_ENCRYPTION_KEY = '';

let identity: typeof import('../identity');
let approvals: typeof import('../state/approvals');
let store: typeof import('../state/store');
let office: typeof import('../office/client');

beforeAll(async () => {
  identity = await import('../identity');
  approvals = await import('../state/approvals');
  store = await import('../state/store');
  office = await import('../office/client');
});

const dm = { userId: 'u1', chatId: 'dm1', chatType: 'p2p' as const };

describe('Feishu admission and memory isolation', () => {
  it('allows group members to mention the bot without a user allowlist entry', () => {
    const base = { senderId: 'u1', chatId: 'g1', chatType: 'group' as const, mentionedBot: true };
    expect(identity.isAuthorizedMessage(base)).toBe(true);
    expect(identity.isAuthorizedMessage({ ...base, mentionedBot: false })).toBe(false);
    expect(identity.isAuthorizedMessage({ ...base, senderId: 'other' })).toBe(true);
    expect(identity.isAuthorizedMessage({ ...base, senderId: '' })).toBe(false);
  });

  it('shares a group thread but isolates direct messages', () => {
    expect(identity.memoryKeys({ userId: 'u1', chatId: 'g1', chatType: 'group' }))
      .toEqual(identity.memoryKeys({ userId: 'u2', chatId: 'g1', chatType: 'group' }));
    expect(identity.memoryKeys(dm)).not.toEqual(identity.memoryKeys({ userId: 'u2', chatId: 'dm1', chatType: 'p2p' }));
    expect(identity.memoryKeys(dm)).not.toEqual(identity.memoryKeys({ userId: 'u1', chatId: 'g1', chatType: 'group' }));
  });

  it('admits any member mentioning the bot in a newly invited group without user or chat ID settings', () => {
    const message = { senderId: 'u1', chatId: 'new-group', chatType: 'group' as const, mentionedBot: true };
    expect(identity.isAuthorizedMessage(message)).toBe(true);
    expect(identity.isAuthorizedMessage({ ...message, senderId: 'other' })).toBe(true);
    expect(identity.isAuthorizedMessage({ ...message, mentionedBot: false })).toBe(false);
    expect(identity.isAuthorizedMessage({ ...message, chatType: 'p2p' })).toBe(false);
    expect(identity.isAuthorizedMessage({ ...message, senderId: 'other', chatType: 'p2p' })).toBe(false);
  });
});

describe('persistent admission and approval', () => {
  it('deduplicates messages durably', async () => {
    expect(await store.markMessageSeen('m1')).toBe(true);
    expect(await store.markMessageSeen('m1')).toBe(false);
  });

  it('allows only the original requester and claims one callback', async () => {
    const action = await approvals.proposeAction(dm, { kind: 'task_create', summary: 'Prepare report' });
    let calls = 0;
    const execute = async () => { calls++; return { task: 't1' }; };
    expect((await approvals.decideAction(action.id, { userId: 'u2', chatId: 'dm1' }, true, execute)).status).toBe('denied');
    expect((await approvals.decideAction(action.id, { userId: 'u1', chatId: 'g1' }, true, execute)).status).toBe('denied');
    expect((await approvals.decideAction(action.id, dm, true, execute)).status).toBe('succeeded');
    expect((await approvals.decideAction(action.id, dm, true, execute)).status).toBe('succeeded');
    expect(calls).toBe(1);
  });

  it('does not execute denied or expired proposals', async () => {
    let calls = 0;
    const execute = async () => { calls++; };
    const denied = await approvals.proposeAction(dm, { kind: 'doc_create', title: 'Draft' });
    expect((await approvals.decideAction(denied.id, dm, false, execute)).status).toBe('denied');
    const expired = await approvals.proposeAction(dm, { kind: 'task_create', summary: 'Old' });
    await (await store.stateDb()).execute({ sql: 'UPDATE actions SET expires_at = 0 WHERE id = ?', args: [expired.id] });
    expect((await approvals.decideAction(expired.id, dm, true, execute)).status).toBe('expired');
    expect(calls).toBe(0);
  });

  it('invalidates in-flight tool approvals after restart and keeps office proposals', async () => {
    const tool = await approvals.proposeAction(dm, { kind: 'mastra_tool', toolName: 'write_file', arguments: {} }, { toolCallId: 'call-1', sessionId: 'dm:u1' });
    const officeAction = await approvals.proposeAction(dm, { kind: 'task_create', summary: 'Keep' });
    await store.recoverInterruptedApprovals();
    expect((await store.getAction(tool.id))?.status).toBe('interrupted');
    expect((await store.getAction(officeAction.id))?.status).toBe('pending');
  });

  it('records API failure and never replays an uncertain write', async () => {
    const action = await approvals.proposeAction(dm, { kind: 'task_create', summary: 'Fail once' });
    let calls = 0;
    const execute = async () => { calls++; throw new Error('upstream timeout'); };
    expect((await approvals.decideAction(action.id, dm, true, execute)).status).toBe('failed');
    expect((await approvals.decideAction(action.id, dm, true, execute)).status).toBe('failed');
    expect(calls).toBe(1);
  });
});

describe('office API transport', () => {
  it('uses app-accessible endpoints and surfaces private API failure', async () => {
    const requests: unknown[] = [];
    const api = new office.OfficeApi({
      request: async request => { requests.push(request); return { code: 0, data: { content: 'notes' } }; },
      send: async () => ({ ok: true }),
    });
    expect(await api.readDocument('doc1')).toEqual({ content: 'notes' });
    await api.execute({ kind: 'calendar_create', calendarId: 'cal1', summary: 'Sync', startTime: '1760000000', endTime: '1760003600' });
    expect(requests).toMatchObject([
      { method: 'GET', url: '/open-apis/docx/v1/documents/doc1/raw_content' },
      { method: 'POST', url: '/open-apis/calendar/v4/calendars/cal1/events' },
    ]);
    const failing = new office.OfficeApi({ request: async () => ({ code: 403, msg: 'not permitted' }), send: async () => undefined });
    await expect(failing.readTask('t1')).rejects.toThrow('Feishu API 403');
    const unavailable = new office.OfficeApi({ request: async () => ({ code: 404, msg: 'unsupported' }), send: async () => undefined });
    await expect(unavailable.readTask('t1')).rejects.toThrow('does not provide');
    await expect(api.execute({ kind: 'message_send', targetChatId: 'outside', text: 'hello' })).rejects.toThrow('not allowlisted');
  });

  it('preserves Feishu error codes from non-2xx HTTP responses', async () => {
    const failure = Object.assign(new Error('Request failed with status code 400'), {
      response: { status: 400, data: { code: 99991679, msg: 'calendar read scope required' } },
    });
    const api = new office.OfficeApi({ request: async () => { throw failure; }, send: async () => undefined });
    await expect(api.primaryCalendar('user-token')).rejects.toThrow('Feishu API 99991679: calendar read scope required');
  });
});

it('does not issue a Harness approval card for an office tool with its own approval flow', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const cards: unknown[] = [];
  const decisions: unknown[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chatId: string, content: unknown) => { cards.push(content); return {}; },
  } as unknown as LarkChannel);
  const session = { respondToToolApproval: async (decision: unknown) => { decisions.push(decision); } };
  await Reflect.get(gateway, 'onToolApproval').call(gateway, {
    toolCallId: 'calendar-call-1', toolName: 'createCalendarTool', args: { summary: 'Meeting', permissionMode: 'user' },
  }, dm, session);
  expect(cards).toHaveLength(0);
  expect(decisions).toEqual([{ toolCallId: 'calendar-call-1', decision: 'approve' }]);
});

it('sends ask_user as a plain question without a misleading success approval card', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const sent: unknown[] = [];
  const decisions: unknown[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chatId: string, payload: unknown) => { sent.push(payload); return {}; },
  } as unknown as LarkChannel);
  const session = { respondToToolApproval: async (decision: unknown) => { decisions.push(decision); } };
  await Reflect.get(gateway, 'onToolApproval').call(gateway, {
    toolCallId: 'ask-1', toolName: 'ask_user', args: { question: '请选择处理方式', options: [{ label: '核对结果' }, { label: '重新发起' }] },
  }, dm, session);
  expect(decisions).toEqual([{ toolCallId: 'ask-1', decision: 'decline' }]);
  expect(sent).toEqual([{ text: '请选择处理方式\n可回复：核对结果、重新发起' }]);
});

it('handles mocked Feishu card callbacks exactly once', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const requests: unknown[] = [];
  const updates: unknown[] = [];
  const channel = {
    rawClient: { request: async (request: unknown) => { requests.push(request); return { code: 0, data: { task: { guid: 'task-1' } } }; } },
    send: async () => ({ messageId: 'card-1' }),
    updateCard: async (_messageId: string, card: unknown) => { updates.push(card); },
  } as unknown as LarkChannel;
  const gateway = new FeishuGateway(channel);
  const action = await approvals.proposeAction(dm, { kind: 'task_create', summary: 'Approved task' });
  const event = (openId: string): CardActionEvent => ({
    messageId: 'card-1', chatId: 'dm1', operator: { openId },
    action: { tag: 'button', value: { approvalId: action.id, decision: 'approve' } },
  });
  await gateway.onCardAction(event('u2'));
  expect(requests).toHaveLength(0);
  await gateway.onCardAction(event('u1'));
  await gateway.onCardAction(event('u1'));
  expect(requests).toHaveLength(1);
  expect(requests[0]).toMatchObject({ data: { client_token: action.id } });
  expect(updates).toHaveLength(2);
});

it('allows an unlisted group member to approve only their own write', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { requestOfficeApproval, setApprovalNotifier } = await import('../runtime/approval-bridge');
  const groupCaller = { userId: 'other', chatId: 'new-group', chatType: 'group' as const };
  const cards: string[] = [];
  const requests: unknown[] = [];
  setApprovalNotifier(async (_chatId, actionId) => { cards.push(actionId); });
  try {
    const proposal = await requestOfficeApproval(groupCaller, { kind: 'task_create', summary: 'Group task' });
    expect(proposal.status).toBe('pending');
    expect(cards).toEqual([proposal.approvalId]);
    const gateway = new FeishuGateway({
      rawClient: { request: async (request: unknown) => { requests.push(request); return { code: 0, data: { task: { guid: 'group-task' } } }; } },
      send: async () => ({ messageId: 'group-card' }),
      updateCard: async () => undefined,
    } as unknown as LarkChannel);
    const callback = (openId: string): CardActionEvent => ({
      messageId: 'group-card', chatId: 'new-group', operator: { openId },
      action: { tag: 'button', value: { approvalId: proposal.approvalId, decision: 'approve' } },
    });
    await gateway.onCardAction(callback('u2'));
    expect(requests).toHaveLength(0);
    await gateway.onCardAction(callback('other'));
    expect(requests).toHaveLength(1);
  } finally {
    setApprovalNotifier(undefined);
  }
});

it('keeps cross-chat message sending unavailable to an unlisted group member', async () => {
  const { requestOfficeApproval, setApprovalNotifier } = await import('../runtime/approval-bridge');
  let cards = 0;
  setApprovalNotifier(async () => { cards++; });
  try {
    await expect(requestOfficeApproval({ userId: 'other', chatId: 'new-group', chatType: 'group' }, {
      kind: 'message_send', targetChatId: 'target', text: 'Outside group',
    })).rejects.toThrow('Cross-chat');
    expect(cards).toBe(0);
  } finally {
    setApprovalNotifier(undefined);
  }
});

it('rejects a cross-chat send card for an unlisted group member even if a proposal already exists', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const sent: string[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (chatId: string) => { sent.push(chatId); return { messageId: 'sent' }; },
    updateCard: async () => undefined,
  } as unknown as LarkChannel);
  const action = await approvals.proposeAction({ userId: 'other', chatId: 'new-group', chatType: 'group' }, {
    kind: 'message_send', targetChatId: 'target', text: 'Outside group',
  });
  await gateway.onCardAction({
    messageId: 'cross-chat-card', chatId: 'new-group', operator: { openId: 'other' },
    action: { tag: 'button', value: { approvalId: action.id, decision: 'approve' } },
  });
  expect(sent).toHaveLength(0);
  expect((await store.getAction(action.id))?.status).toBe('pending');
});

it('does not approve a write from an unlisted direct chat that has a group-shaped chat ID', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const requests: unknown[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async (request: unknown) => { requests.push(request); return { code: 0, data: {} }; } },
    send: async () => ({ messageId: 'dm-card' }),
    updateCard: async () => undefined,
  } as unknown as LarkChannel);
  const action = await approvals.proposeAction({ userId: 'u1', chatId: 'oc_unlisted_dm', chatType: 'p2p' }, { kind: 'task_create', summary: 'Denied task' });
  await gateway.onCardAction({
    messageId: 'dm-card', chatId: 'oc_unlisted_dm', operator: { openId: 'u1' },
    action: { tag: 'button', value: { approvalId: action.id, decision: 'approve' } },
  });
  expect(requests).toHaveLength(0);
  expect((await store.getAction(action.id))?.status).toBe('pending');
});

it('shows the created calendar date and event link after card approval', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const cards: unknown[] = [];
  const channel = {
    rawClient: { request: async () => ({ code: 0, data: { event: {
      event_id: 'event-1', app_link: 'https://applink.feishu.cn/client/calendar/event/detail?key=event-1',
    } } }) },
    send: async () => ({ messageId: 'calendar-card' }),
    updateCard: async (_messageId: string, card: unknown) => { cards.push(card); },
  } as unknown as LarkChannel;
  const gateway = new FeishuGateway(channel);
  const action = await approvals.proposeAction(dm, {
    kind: 'calendar_create', calendarId: 'shared-calendar', permissionMode: 'app', summary: '10月工作总结大会',
    startTime: '1793062800', endTime: '1793066400',
  });
  await gateway.onCardAction({
    messageId: 'calendar-card', chatId: 'dm1', operator: { openId: 'u1' },
    action: { tag: 'button', value: { approvalId: action.id, decision: 'approve' } },
  });
  expect(JSON.stringify(cards[0])).toContain('2026-10-27 09:00');
  expect(JSON.stringify(cards[0])).toContain('https://applink.feishu.cn/client/calendar/event/detail?key=event-1');
});

it('reveals only the sender and current chat IDs before allowlisting', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const sent: Array<{ chatId: string; text: string }> = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (chatId: string, input: { text: string }) => { sent.push({ chatId, text: input.text }); return {}; },
  } as unknown as LarkChannel);
  const message = (messageId: string, overrides: Partial<NormalizedMessage>): NormalizedMessage => ({
    messageId, senderId: 'ou_new_user', chatId: 'oc_new_dm', chatType: 'p2p',
    content: '/id', mentionedBot: false, rawContentType: 'text', resources: [], mentions: [],
    mentionAll: false, createTime: Date.now(), ...overrides,
  });
  await gateway.onMessage(message('discover-dm', {}));
  expect(sent).toEqual([{ chatId: 'oc_new_dm', text: expect.stringContaining('ou_new_user') }]);
  expect(sent[0]?.text).toContain('oc_new_dm');
  await gateway.onMessage(message('discover-group', { chatId: 'oc_new_group', chatType: 'group', mentionedBot: true }));
  expect(sent[1]).toEqual({ chatId: 'oc_new_group', text: expect.stringContaining('oc_new_group') });
  expect(sent[1]?.text).not.toContain('ou_new_user');
  await gateway.onMessage(message('discover-no-mention', { chatId: 'oc_new_group', chatType: 'group' }));
  await gateway.onMessage(message('discover-other-command', { content: 'create a task' }));
  expect(sent).toHaveLength(2);
});

it('restarts a failed release only from the original requester’s mentioned group reply without replaying the failed approval', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const { setOfficeApi, setOfficeTextSender } = await import('../runtime/office-bridge');
  const { setApprovalNotifier } = await import('../runtime/approval-bridge');
  const sent: unknown[] = [];
  const channel = {
    rawClient: { request: async (request: { url: string; params?: Record<string, unknown> }) => {
      if (request.url === '/open-apis/drive/v1/files') return { code: 0, data: { files: request.params?.folder_token === 'template-retry' ? [{ token: 'doxcnRetry', name: '模板', type: 'docx', modified_time: '1' }] : [], has_more: false } };
      if (request.url.endsWith('/members')) return { code: 0, data: { items: [{ member_type: 'openchat', member_id: 'g1', perm: 'edit' }] } };
      if (request.url.endsWith('/blocks')) return { code: 0, data: { items: [{ block_id: 'blockRetry', text: { elements: [{ text_run: { content: '{{负责人}}' } }] } }], has_more: false } };
      throw new Error(`Unexpected Feishu request ${request.url}`);
    } },
    send: async (_chatId: string, payload: unknown) => { sent.push(payload); return {}; },
  } as unknown as LarkChannel;
  const gateway = new FeishuGateway(channel);
  const requester = { userId: 'u1', chatId: 'g1', chatType: 'group' as const };
  const service = new ReleaseMaterialsService(gateway.office);
  const manifest = await service.prepare({
    templateFolderUrl: 'https://feishu.cn/drive/folder/template-retry',
    materialsParentUrl: 'https://feishu.cn/drive/folder/parent-retry',
    releaseId: '202610291638', fields: { 负责人: 'Topson' },
  }, requester);
  const oldJob = await service.createJob(requester, manifest, 1000);
  const oldApproval = await approvals.proposeAction(requester, { kind: 'release_materials', jobId: oldJob.jobId, releaseId: manifest.input.releaseId, fileCount: 1, fieldSummary: '负责人=Topson' });
  const db = await store.stateDb();
  await db.execute({ sql: "UPDATE release_jobs SET status = 'failed', approval_id = ? WHERE id = ?", args: [oldApproval.id, oldJob.jobId] });
  await db.execute({ sql: "UPDATE actions SET status = 'failed' WHERE id = ?", args: [oldApproval.id] });
  setOfficeApi(gateway.office);
  setOfficeTextSender(async (_chatId, text) => { sent.push({ markdown: text }); });
  setApprovalNotifier(async (_chatId, id) => { sent.push({ cardApprovalId: id }); });
  try {
    await gateway.onMessage({ messageId: 'release-retry-plain', senderId: 'u1', chatId: 'g1', chatType: 'group', content: '重新发起', mentionedBot: false, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: 2000 });
    expect((await db.execute({ sql: "SELECT id FROM release_jobs WHERE requester_id = 'u1' AND chat_id = 'g1'" })).rows).toHaveLength(1);
    expect(sent).toHaveLength(0);
    await gateway.onMessage({ messageId: 'release-retry-other', senderId: 'u2', chatId: 'g1', chatType: 'group', content: '重新发起', mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: 2000 });
    expect(JSON.stringify(sent)).toContain('没有属于你的可重试下发批次');
    sent.length = 0;
    await gateway.onMessage({ messageId: 'release-retry-mentioned', senderId: 'u1', chatId: 'g1', chatType: 'group', content: '重新发起', mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: 2001 });
    const jobs = (await db.execute({ sql: "SELECT status, approval_id FROM release_jobs WHERE requester_id = 'u1' AND chat_id = 'g1' ORDER BY created_at DESC" })).rows;
    expect(jobs).toHaveLength(2);
    expect(jobs[0].status).toBe('pending');
    expect(jobs[0].approval_id).not.toBe(oldApproval.id);
    expect(sent.filter(item => 'cardApprovalId' in (item as object))).toHaveLength(1);
    expect((await store.getAction(oldApproval.id))?.status).toBe('failed');
  } finally {
    setOfficeApi(undefined);
    setOfficeTextSender(undefined);
    setApprovalNotifier(undefined);
  }
});

it('rechecks a completed release in Drive and starts a fresh approval if its folder was deleted', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const { setOfficeApi, setOfficeTextSender } = await import('../runtime/office-bridge');
  const { setApprovalNotifier } = await import('../runtime/approval-bridge');
  const sent: unknown[] = [];
  const caller = { userId: 'u-completed', chatId: 'g-completed', chatType: 'group' as const };
  const channel = {
    rawClient: { request: async (request: { url: string; params?: Record<string, unknown> }) => {
      if (request.url === '/open-apis/drive/v1/files') return { code: 0, data: {
        files: request.params?.folder_token === 'template-completed' ? [{ token: 'doxcnCompleted', name: '模板', type: 'docx' }] : [],
        has_more: false,
      } };
      if (request.url.endsWith('/members')) return { code: 0, data: { items: [{ member_type: 'openchat', member_id: caller.chatId, perm: 'edit' }] } };
      if (request.url.endsWith('/blocks')) return { code: 0, data: { items: [{ block_id: 'body', block_type: 2, text: { elements: [{ text_run: { content: '{{负责人}}' } }] } }], has_more: false } };
      throw new Error(`Unexpected request ${request.url}`);
    } },
    send: async (_chatId: string, payload: unknown) => { sent.push(payload); return {}; },
  } as unknown as LarkChannel;
  const gateway = new FeishuGateway(channel);
  const input = { templateFolderUrl: 'https://feishu.cn/drive/folder/template-completed',
    materialsParentUrl: 'https://feishu.cn/drive/folder/parent-completed', releaseId: '202610291638', fields: { 负责人: 'Topson' } };
  const service = new ReleaseMaterialsService(gateway.office);
  const old = await service.createJob(caller, await service.prepare(input), 1000);
  const db = await store.stateDb();
  await db.execute({ sql: "UPDATE release_jobs SET status = 'completed', folder_token = 'deleted-folder' WHERE id = ?", args: [old.jobId] });
  setOfficeApi(gateway.office);
  setOfficeTextSender(async (_chatId, text) => { sent.push({ markdown: text }); });
  setApprovalNotifier(async (_chatId, id) => { sent.push({ cardApprovalId: id }); });
  try {
    const handled = await Reflect.get(gateway, 'handleKnownReleaseRequest').call(gateway,
      `请再次下发 ${input.releaseId} ${input.templateFolderUrl} ${input.materialsParentUrl} Topson`, caller, 'message-1', 2000);
    expect(handled).toBe(true);
    expect((await db.execute({ sql: 'SELECT status FROM release_jobs WHERE id = ?', args: [old.jobId] })).rows[0].status).toBe('failed');
    const jobs = (await db.execute({ sql: 'SELECT status FROM release_jobs WHERE requester_id = ? ORDER BY created_at DESC', args: [caller.userId] })).rows;
    expect(jobs.map(row => row.status)).toContain('pending');
    expect(sent.filter(item => 'cardApprovalId' in (item as object))).toHaveLength(1);
    sent.length = 0;
    await Reflect.get(gateway, 'handleKnownReleaseRequest').call(gateway,
      `再次下发 ${input.releaseId} ${input.templateFolderUrl} ${input.materialsParentUrl} Topson`, caller, 'message-2', 3000);
    expect(JSON.stringify(sent)).toContain('仍在等待确认');
    expect(sent.filter(item => 'cardApprovalId' in (item as object))).toHaveLength(0);
  } finally {
    setOfficeApi(undefined);
    setOfficeTextSender(undefined);
    setApprovalNotifier(undefined);
  }
});

it('explains a release retry preflight failure instead of silently dropping the reply or sending another approval', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const { setOfficeApi, setOfficeTextSender } = await import('../runtime/office-bridge');
  const { setApprovalNotifier } = await import('../runtime/approval-bridge');
  const sent: unknown[] = [];
  const channel = {
    rawClient: { request: async (request: { url: string; params?: Record<string, unknown> }) => {
      if (request.url === '/open-apis/drive/v1/files') return { code: 0, data: { files: request.params?.folder_token === 'template-view' ? [{ token: 'doxcnView', name: '模板', type: 'docx' }] : [], has_more: false } };
      if (request.url.endsWith('/members')) return { code: 0, data: { items: [{ member_type: 'openchat', member_id: 'g1', perm: 'view' }] } };
      throw new Error(`Unexpected Feishu request ${request.url}`);
    } },
    send: async (_chatId: string, payload: unknown) => { sent.push(payload); return {}; },
  } as unknown as LarkChannel;
  const gateway = new FeishuGateway(channel);
  const requester = { userId: 'u1', chatId: 'g1', chatType: 'group' as const };
  const service = new ReleaseMaterialsService(gateway.office);
  const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template-view', materialsParentUrl: 'https://feishu.cn/drive/folder/parent-view', releaseId: '202610291639', fields: { 负责人: 'Topson' } }, templateFolderToken: 'template-view', materialsParentToken: 'parent-view', files: [], ignoredFolders: [] };
  const oldJob = await service.createJob(requester, manifest, 3000);
  const db = await store.stateDb();
  await db.execute({ sql: "UPDATE release_jobs SET status = 'failed' WHERE id = ?", args: [oldJob.jobId] });
  setOfficeApi(gateway.office);
  setOfficeTextSender(async (_chatId, text) => { sent.push({ markdown: text }); });
  setApprovalNotifier(async (_chatId, id) => { sent.push({ cardApprovalId: id }); });
  try {
    await gateway.onMessage({ messageId: 'release-retry-view', senderId: 'u1', chatId: 'g1', chatType: 'group', content: '重新发起', mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: 4000 });
    expect(JSON.stringify(sent)).toContain('仅有查看权限');
    expect(sent.filter(item => 'cardApprovalId' in (item as object))).toHaveLength(0);
    const jobs = (await db.execute({ sql: "SELECT id FROM release_jobs WHERE requester_id = 'u1' AND chat_id = 'g1' AND manifest_json LIKE '%202610291639%'" })).rows;
    expect(jobs).toHaveLength(1);
  } finally {
    setOfficeApi(undefined);
    setOfficeTextSender(undefined);
    setApprovalNotifier(undefined);
  }
});

it('does not relay an Agent claim that a failed release approval is still pending', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const { memoryKeys } = await import('../identity');
  const sent: string[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chatId: string, payload: { markdown?: string; text?: string }) => { sent.push(payload.markdown ?? payload.text ?? ''); return {}; },
  } as unknown as LarkChannel);
  const requester = { userId: 'u1', chatId: 'g1', chatType: 'group' as const };
  const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template-stale', materialsParentUrl: 'https://feishu.cn/drive/folder/parent-stale', releaseId: '202610291640', fields: { 负责人: 'Topson' } }, templateFolderToken: 'template-stale', materialsParentToken: 'parent-stale', files: [], ignoredFolders: [] };
  const service = new ReleaseMaterialsService(gateway.office);
  const job = await service.createJob(requester, manifest, 5000);
  const action = await approvals.proposeAction(requester, { kind: 'release_materials', jobId: job.jobId, releaseId: manifest.input.releaseId, fileCount: 1, fieldSummary: '负责人=Topson' });
  const db = await store.stateDb();
  await db.execute({ sql: "UPDATE release_jobs SET status = 'failed', approval_id = ? WHERE id = ?", args: [action.id, job.jobId] });
  await db.execute({ sql: "UPDATE actions SET status = 'failed' WHERE id = ?", args: [action.id] });
  let subscriber: (event: unknown) => void = () => undefined;
  Reflect.get(gateway, 'sessions').set(memoryKeys(requester).sessionId, {
    subscribe: (callback: (event: unknown) => void) => { subscriber = callback; return () => undefined; },
    sendMessage: async () => {
      subscriber({ type: 'message_update', event: { type: 'text-delta', delta: `这批下发已在等你审批。待审批批次：${action.id}` } });
      subscriber({ type: 'agent_end', reason: 'completed' });
    },
  });
  await Reflect.get(gateway, 'runAgent').call(gateway, '重新发送 202610291640 的下发材料请求', requester, 'release-stale');
  await vi.waitFor(() => expect(sent).toHaveLength(1));
  expect(sent[0]).toContain('此前批次已失败');
  expect(sent[0]).not.toContain(action.id);
});

it('grounds a new release reply in its own pending job instead of historical approvals in Agent memory', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const { memoryKeys } = await import('../identity');
  const sent: string[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chatId: string, payload: { markdown?: string; text?: string }) => { sent.push(payload.markdown ?? payload.text ?? ''); return {}; },
  } as unknown as LarkChannel);
  const requester = { userId: 'u-ground', chatId: 'g-ground', chatType: 'group' as const };
  const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template-ground', materialsParentUrl: 'https://feishu.cn/drive/folder/parent-ground', releaseId: '202610300856', fields: { 版本负责人: 'Ana', 下发日期: '20261031' } }, templateFolderToken: 'template-ground', materialsParentToken: 'parent-ground', files: [], ignoredFolders: [] };
  const job = await new ReleaseMaterialsService(gateway.office).createJob(requester, manifest, 1793321832959);
  const action = await approvals.proposeAction(requester, { kind: 'release_materials', jobId: job.jobId, releaseId: '202610300856', fileCount: 1, fieldSummary: 'Ana' });
  await (await store.stateDb()).execute({ sql: 'UPDATE release_jobs SET approval_id = ? WHERE id = ?', args: [action.id, job.jobId] });
  let subscriber: (event: unknown) => void = () => undefined;
  Reflect.get(gateway, 'sessions').set(memoryKeys(requester).sessionId, {
    subscribe: (callback: (event: unknown) => void) => { subscriber = callback; return () => undefined; },
    sendMessage: async () => {
      subscriber({ type: 'message_update', event: { type: 'text-delta', delta: '你有两个请求；旧批次 202610291638 的审批 87d90d82-72be-41be-a4b3-2535867a2576 仍待确认。两张卡片都等待审批。' } });
      subscriber({ type: 'agent_end', reason: 'completed' });
    },
  });
  await Reflect.get(gateway, 'runAgent').call(gateway, '新增下发版本编号为“202610300856”，版本负责人 Ana，下发日期 20261031', requester, 'new-release-message');
  await vi.waitFor(() => expect(sent).toHaveLength(1));
  expect(sent[0]).toContain('202610300856');
  expect(sent[0]).toContain(action.id);
  expect(sent[0]).not.toContain('202610291638');
  expect(sent[0]).not.toContain('两张卡片');
});

it('reports the completed release when the requester follows up with confirmation', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const sent: string[] = [];
  const gateway = new FeishuGateway({
    rawClient: { request: async () => ({ code: 0, data: {} }) },
    send: async (_chatId: string, payload: { markdown?: string; text?: string }) => { sent.push(payload.markdown ?? payload.text ?? ''); return {}; },
  } as unknown as LarkChannel);
  const requester = { userId: 'u-confirm', chatId: 'g-confirm', chatType: 'group' as const };
  const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template-confirm', materialsParentUrl: 'https://feishu.cn/drive/folder/parent-confirm', releaseId: '202610300857', fields: { 版本负责人: 'Ana' } }, templateFolderToken: 'template-confirm', materialsParentToken: 'parent-confirm', files: [], ignoredFolders: [] };
  const job = await new ReleaseMaterialsService(gateway.office).createJob(requester, manifest, 1793321832959);
  await (await store.stateDb()).execute({ sql: "UPDATE release_jobs SET status = 'completed', folder_url = ?, result_json = ?, updated_at = ? WHERE id = ?", args: ['https://feishu.cn/drive/folder/finished-confirm', JSON.stringify({ folderUrl: 'https://feishu.cn/drive/folder/finished-confirm', files: [{ name: '清单', url: 'https://feishu.cn/sheets/sheet-confirm' }] }), Date.now(), job.jobId] });
  await gateway.onMessage({ messageId: 'confirmed-release-followup', senderId: requester.userId, chatId: requester.chatId, chatType: 'group', content: '通过新批次的审批卡片', mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: Date.now() });
  expect(sent).toHaveLength(1);
  expect(sent[0]).toContain('202610300857');
  expect(sent[0]).toContain('https://feishu.cn/drive/folder/finished-confirm');
  expect(sent[0]).not.toContain('待审批');
});

it('expires an overdue release approval before deciding whether a retry may start', async () => {
  const { FeishuGateway } = await import('../gateway/gateway');
  const { ReleaseMaterialsService } = await import('../office/release-materials');
  const gateway = new FeishuGateway({ rawClient: { request: async () => ({ code: 0, data: {} }) }, send: async () => ({}) } as unknown as LarkChannel);
  const requester = { userId: 'u1', chatId: 'g1', chatType: 'group' as const };
  const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template-expired', materialsParentUrl: 'https://feishu.cn/drive/folder/parent-expired', releaseId: '202610291641', fields: { 负责人: 'Topson' } }, templateFolderToken: 'template-expired', materialsParentToken: 'parent-expired', files: [], ignoredFolders: [] };
  const job = await new ReleaseMaterialsService(gateway.office).createJob(requester, manifest, 6000);
  const action = await approvals.proposeAction(requester, { kind: 'release_materials', jobId: job.jobId, releaseId: manifest.input.releaseId, fileCount: 1, fieldSummary: '负责人=Topson' });
  const db = await store.stateDb();
  await db.execute({ sql: 'UPDATE release_jobs SET approval_id = ? WHERE id = ?', args: [action.id, job.jobId] });
  await db.execute({ sql: 'UPDATE actions SET expires_at = 0 WHERE id = ?', args: [action.id] });
  const latest = await Reflect.get(gateway, 'latestReleaseJob').call(gateway, requester);
  expect(latest.status).toBe('expired');
  expect((await store.getAction(action.id))?.status).toBe('expired');
  expect((await db.execute({ sql: 'SELECT status FROM release_jobs WHERE id = ?', args: [job.jobId] })).rows[0].status).toBe('expired');
});

it('asks for calendar identity before approval, persists the choice, and requires an app calendar ID', async () => {
  const { requestCalendarPermission, resolvePermissionReply, setPermissionNotifier } = await import('../office/permission-choice');
  const { setApprovalNotifier } = await import('../runtime/approval-bridge');
  const questions: string[] = [];
  const cards: string[] = [];
  setPermissionNotifier(async (_chatId, text) => { questions.push(text); });
  setApprovalNotifier(async (_chatId, _id, summary) => { cards.push(summary); });
  try {
    const result = await requestCalendarPermission(dm, {
      kind: 'calendar_create', permissionMode: 'user', summary: '我的日程', startTime: '1760000000', endTime: '1760003600',
    });
    expect(result.status).toBe('permission_choice_required');
    expect(questions).toHaveLength(1);
    expect(cards).toHaveLength(0);
    expect(await resolvePermissionReply({ userId: 'u2', chatId: 'dm1', chatType: 'p2p' }, '应用权限')).toBe(false);
    expect(await resolvePermissionReply(dm, '应用权限')).toBe(true);
    expect(questions.at(-1)).toContain('共享日历 ID');
    expect(cards).toHaveLength(0);
    expect(await resolvePermissionReply(dm, 'cal_123')).toBe(true);
    expect(cards).toHaveLength(1);
    expect(cards[0]).toContain('应用身份');
    expect(await resolvePermissionReply(dm, 'cal_123')).toBe(false);
  } finally {
    setPermissionNotifier(undefined);
    setApprovalNotifier(undefined);
  }
});

it('never uses the app token when an approved calendar action requires user identity', async () => {
  const requests: unknown[] = [];
  const api = new office.OfficeApi({ request: async request => { requests.push(request); return { code: 0, data: {} }; }, send: async () => undefined });
  const action = { kind: 'calendar_create' as const, permissionMode: 'user' as const, calendarId: 'personal', summary: 'Mine', startTime: '1760000000', endTime: '1760003600' };
  await expect(api.execute(action)).rejects.toThrow('缺少请求者');
  expect(requests).toHaveLength(0);
  await api.execute(action, 'key', 'u-token');
  expect(requests).toMatchObject([{ userAccessToken: 'u-token' }]);
});

it('explains missing OAuth configuration in chat without generating an approval card', async () => {
  const { requestCalendarPermission, resolvePermissionReply, setPermissionNotifier } = await import('../office/permission-choice');
  const { setApprovalNotifier } = await import('../runtime/approval-bridge');
  const messages: string[] = [];
  let approvals = 0;
  setPermissionNotifier(async (_chat, text) => { messages.push(text); });
  setApprovalNotifier(async () => { approvals++; });
  try {
    await requestCalendarPermission(dm, { kind: 'calendar_create', permissionMode: 'user', summary: 'Lunch', startTime: '1760000000', endTime: '1760003600' });
    expect(await resolvePermissionReply(dm, '用户权限')).toBe(true);
    expect(messages.at(-1)).toContain('FEISHU_OAUTH_AUTHORIZE_URL');
    expect(approvals).toBe(0);
    expect(await resolvePermissionReply(dm, '取消')).toBe(true);
  } finally {
    setPermissionNotifier(undefined);
    setApprovalNotifier(undefined);
  }
});
