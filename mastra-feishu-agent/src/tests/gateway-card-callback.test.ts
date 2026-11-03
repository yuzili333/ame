import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { LarkChannel, RawCardActionEvent } from '@larksuiteoapi/node-sdk';

const dataDir = mkdtempSync(join(tmpdir(), 'feishu-card-callback-test-'));
let FeishuGateway: typeof import('../gateway/gateway').FeishuGateway;
let proposeAction: typeof import('../state/approvals').proposeAction;
let stateDb: typeof import('../state/store').stateDb;

beforeAll(async () => {
  vi.stubEnv('AGENT_DATA_DIR', dataDir);
  vi.stubEnv('FEISHU_ALLOWED_USERS', 'ou_user');
  vi.stubEnv('FEISHU_ALLOWED_CHATS', 'oc_chat');
  vi.stubEnv('FEISHU_OAUTH_REDIRECT_URI', '');
  vi.resetModules();
  ({ FeishuGateway } = await import('../gateway/gateway'));
  ({ proposeAction } = await import('../state/approvals'));
  ({ stateDb } = await import('../state/store'));
});
afterAll(() => { vi.unstubAllEnvs(); rmSync(dataDir, { recursive: true, force: true }); });

function rawAction(approvalId: string, openId = 'ou_user', chatId = 'oc_chat'): RawCardActionEvent {
  return { context: { open_message_id: 'om_card', open_chat_id: chatId }, operator: { open_id: openId },
    action: { value: { approvalId, decision: 'approve' } } } as RawCardActionEvent;
}

describe('private Feishu card callback acknowledgement', () => {
  it('shows each release file and its current state in the execution card', async () => {
    const { releaseProgressCard } = await import('../gateway/cards');
    const card = releaseProgressCard('executing', '202611030909', [
      { name: '下发步骤', status: 'completed' },
      { name: '回退步骤', status: 'writing' },
      { name: '检查单', status: 'pending' },
    ]);
    const text = JSON.stringify(card);
    expect(text).toContain('202611030909');
    expect(text).toContain('下发步骤：更新成功');
    expect(text).toContain('回退步骤：更新中');
    expect(text).toContain('检查单：待更新');
    const failed = JSON.stringify(releaseProgressCard('failed', '202611030909', [
      { name: '下发步骤', status: 'completed' },
      { name: '回退步骤', status: 'uncertain' },
      { name: '检查单', status: 'pending' },
    ]));
    expect(failed).toContain('下发步骤：更新成功');
    expect(failed).toContain('回退步骤：更新失败（结果需核对）');
    expect(failed).toContain('检查单：未执行');
  });

  it('returns the release file list in the immediate card callback', async () => {
    const { ReleaseMaterialsService } = await import('../office/release-materials');
    const channel = { rawClient: {}, updateCard: async () => undefined } as unknown as LarkChannel;
    const gateway = new FeishuGateway(channel);
    const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template', materialsParentUrl: 'https://feishu.cn/drive/folder/parent', releaseId: '202611030909', fields: {} }, templateFolderToken: 'template', materialsParentToken: 'parent', ignoredFolders: [], files: [
      { token: 'doc1', name: '下发步骤', type: 'docx', patches: [], fingerprint: 'one' },
      { token: 'sheet1', name: '检查单', type: 'sheet', patches: [], fingerprint: 'two' },
    ] };
    const caller = { userId: 'ou_user', chatId: 'oc_chat', chatType: 'p2p' as const };
    const { jobId } = await new ReleaseMaterialsService(gateway.office).createJob(caller, manifest);
    const action = await proposeAction(caller, { kind: 'release_materials', jobId, releaseId: '202611030909', fileCount: 2, fieldSummary: '无' });
    await (await stateDb()).execute({ sql: 'UPDATE actions SET card_message_id = ? WHERE id = ?', args: ['om_card', action.id] });
    Reflect.set(gateway, 'onCardAction', async () => undefined);
    const info = vi.spyOn(console, 'info').mockImplementation(() => {});
    try {
      const response = await gateway.ackCardAction(rawAction(action.id));
      const card = JSON.stringify(response);
      expect(card).toContain('下发步骤：待更新');
      expect(card).toContain('检查单：待更新');
    } finally { info.mockRestore(); }
  });

  it('updates the same release card as persisted file states change', async () => {
    const { ReleaseMaterialsService } = await import('../office/release-materials');
    const updates: string[] = [];
    const gateway = new FeishuGateway({ rawClient: {}, updateCard: async (_id: string, card: unknown) => { updates.push(JSON.stringify(card)); } } as unknown as LarkChannel);
    const manifest = { input: { templateFolderUrl: 'https://feishu.cn/drive/folder/template', materialsParentUrl: 'https://feishu.cn/drive/folder/parent', releaseId: '202611030910', fields: {} }, templateFolderToken: 'template', materialsParentToken: 'parent', ignoredFolders: [], files: [
      { token: 'doc2', name: '回退步骤', type: 'docx', patches: [], fingerprint: 'one' },
    ] };
    const { jobId } = await new ReleaseMaterialsService(gateway.office).createJob({ userId: 'ou_user', chatId: 'oc_chat', chatType: 'p2p' }, manifest);
    const db = await stateDb();
    const stop = Reflect.get(gateway, 'watchReleaseProgress').call(gateway, jobId, 'om_card') as () => Promise<void>;
    try {
      await db.execute({ sql: "UPDATE release_files SET status = 'writing' WHERE job_id = ?", args: [jobId] });
      await vi.waitFor(() => expect(updates.some(card => card.includes('回退步骤：更新中'))).toBe(true), { timeout: 3000 });
      await db.execute({ sql: "UPDATE release_files SET status = 'completed' WHERE job_id = ?", args: [jobId] });
      await vi.waitFor(() => expect(updates.some(card => card.includes('回退步骤：更新成功'))).toBe(true), { timeout: 3000 });
    } finally { await stop(); }
  });

  it('keeps the card callback installed after channel.connect registers SDK defaults', async () => {
    const handlers = new Map<string, (event: RawCardActionEvent) => unknown>();
    const requests: unknown[] = [];
    const dispatcher = {
      register: (items: Record<string, (event: RawCardActionEvent) => unknown>) => {
        for (const [name, handler] of Object.entries(items)) handlers.set(name, handler);
      },
      unregister: (...names: string[]) => { for (const name of names) handlers.delete(name); },
    };
    const channel = {
      dispatcher,
      rawClient: { request: async (request: unknown) => { requests.push(request); return { code: 0, data: { task: { guid: 'created-task' } } }; } },
      on: () => undefined,
      connect: async () => { handlers.set('card.action.trigger', async () => undefined); },
      disconnect: async () => undefined,
      updateCard: async () => undefined,
    } as unknown as LarkChannel;
    const action = await proposeAction({ userId: 'ou_other', chatId: 'oc_group', chatType: 'group' }, { kind: 'task_create', summary: 'test' });
    const gateway = new FeishuGateway(channel);
    await gateway.start();
    try {
      const response = await handlers.get('card.action.trigger')?.(rawAction(action.id, 'ou_other', 'oc_group'));
      expect(response).toMatchObject({ card: { type: 'raw', data: { schema: '2.0' } } });
      await vi.waitFor(async () => {
        const row = (await (await stateDb()).execute({ sql: 'SELECT status FROM actions WHERE id = ?', args: [action.id] })).rows[0];
        expect(row?.status).toBe('succeeded');
      });
      expect(requests).toHaveLength(1);
    } finally { await gateway.stop(); }
  });

  it('returns a valid card before a release workflow finishes', async () => {
    const action = await proposeAction({ userId: 'ou_user', chatId: 'oc_chat', chatType: 'p2p' }, { kind: 'task_create', summary: 'test' });
    await (await stateDb()).execute({ sql: 'UPDATE actions SET card_message_id = ? WHERE id = ?', args: ['om_card', action.id] });
    const gateway = new FeishuGateway({ rawClient: {} } as LarkChannel);
    let finish!: () => void;
    const processing = new Promise<void>(resolve => { finish = resolve; });
    const handle = vi.fn(async () => { await processing; });
    Reflect.set(gateway, 'onCardAction', handle);
    const info = vi.spyOn(console, 'info').mockImplementation(() => {});
    const response = await gateway.ackCardAction(rawAction(action.id));
    expect(response).toMatchObject({ card: { type: 'raw', data: { schema: '2.0', header: { title: { content: '操作状态' } } } } });
    expect(JSON.stringify(response)).toContain('正在执行');
    expect(handle).toHaveBeenCalledTimes(1);
    finish();
    await processing;
    info.mockRestore();
  });

  it('keeps an unauthorized click from starting work or disabling the requester card', async () => {
    const action = await proposeAction({ userId: 'ou_user', chatId: 'oc_chat', chatType: 'p2p' }, { kind: 'task_create', summary: 'test' });
    const gateway = new FeishuGateway({ rawClient: {} } as LarkChannel);
    const handle = vi.fn();
    Reflect.set(gateway, 'onCardAction', handle);
    const response = await gateway.ackCardAction(rawAction(action.id, 'ou_other'));
    expect(response).toMatchObject({ toast: { type: 'error' } });
    expect(response).not.toHaveProperty('card');
    expect(handle).not.toHaveBeenCalled();
  });

  it('acknowledges an authorized group member without a user allowlist entry', async () => {
    const action = await proposeAction({ userId: 'ou_other', chatId: 'oc_group', chatType: 'group' }, { kind: 'task_create', summary: 'group test' });
    const gateway = new FeishuGateway({ rawClient: {} } as LarkChannel);
    const handle = vi.fn(async () => undefined);
    Reflect.set(gateway, 'onCardAction', handle);
    const info = vi.spyOn(console, 'info').mockImplementation(() => {});
    const response = await gateway.ackCardAction(rawAction(action.id, 'ou_other', 'oc_group'));
    expect(response).toMatchObject({ card: { type: 'raw', data: { schema: '2.0', header: { title: { content: '操作状态' } } } } });
    expect(handle).toHaveBeenCalledTimes(1);
    info.mockRestore();
  });

  it('reports an incomplete callback without starting approval execution', async () => {
    const gateway = new FeishuGateway({ rawClient: {} } as LarkChannel);
    const handle = vi.fn();
    Reflect.set(gateway, 'onCardAction', handle);
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => {});
    expect(await gateway.ackCardAction({} as RawCardActionEvent)).toMatchObject({ toast: { type: 'error' } });
    expect(handle).not.toHaveBeenCalled();
    warning.mockRestore();
  });
});
