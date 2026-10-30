import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { LarkChannel, NormalizedMessage } from '@larksuiteoapi/node-sdk';

const dataDir = mkdtempSync(join(tmpdir(), 'feishu-recall-test-'));
beforeAll(() => { vi.stubEnv('AGENT_DATA_DIR', dataDir); vi.resetModules(); });
afterAll(() => { vi.unstubAllEnvs(); rmSync(dataDir, { recursive: true, force: true }); });

function message(messageId: string): NormalizedMessage {
  return { messageId, senderId: 'u-recall', chatId: 'g-recall', chatType: 'group', content: '新增下发任务',
    mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: Date.now() };
}

describe('recalled requests', () => {
  it('invalidates the original pending approval and prevents a replayed request from running', async () => {
    const { FeishuGateway } = await import('../gateway/gateway');
    const { proposeAction, decideAction } = await import('../state/approvals');
    const { getAction, stateDb } = await import('../state/store');
    const sent: unknown[] = [];
    const updated: string[] = [];
    const gateway = new FeishuGateway({ send: async (_chat: string, payload: unknown) => { sent.push(payload); return { messageId: 'bot-card' }; },
      updateCard: async (id: string) => { updated.push(id); }, rawClient: { request: async () => ({ code: 0, data: {} }) } } as unknown as LarkChannel);
    const caller = { userId: 'u-recall', chatId: 'g-recall', chatType: 'group' as const, messageId: 'm-recalled' };
    const action = await proposeAction(caller, { kind: 'task_create', summary: 'old task' });
    await (await stateDb()).execute({ sql: 'UPDATE actions SET card_message_id = ? WHERE id = ?', args: ['bot-card', action.id] });
    await gateway.onMessageRecalled({ messageId: 'm-recalled', chatId: 'g-recall' });
    expect((await getAction(action.id))?.status).toBe('denied');
    expect(updated).toEqual(['bot-card']);
    const decision = await decideAction(action.id, caller, true, async () => { throw new Error('must not write'); });
    expect(decision.status).toBe('denied');
    await gateway.onMessage(message('m-recalled'));
    expect(sent).toHaveLength(0);
  });

  it('cancels a queued request by message ID without affecting another queued request', async () => {
    const { FeishuGateway } = await import('../gateway/gateway');
    const { stateDb } = await import('../state/store');
    const sent: unknown[] = [];
    const gateway = new FeishuGateway({ send: async (_chat: string, payload: unknown) => { sent.push(payload); return { messageId: 'reply' }; },
      rawClient: { request: async () => ({ code: 0, data: {} }) } } as unknown as LarkChannel);
    const cancel = vi.fn();
    Reflect.get(gateway, 'activeRuns').set('group:g-recall', { messageId: 'active', chatId: 'g-recall', cancel });
    await gateway.onMessage(message('queued-1'));
    await gateway.onMessage(message('queued-2'));
    await gateway.onMessageRecalled({ messageId: 'queued-1', chatId: 'g-recall' });
    const rows = (await (await stateDb()).execute({ sql: 'SELECT message_id, status FROM queued_messages ORDER BY created_at' })).rows;
    expect(rows.map(row => [row.message_id, row.status])).toEqual([['queued-1', 'cancelled'], ['queued-2', 'pending']]);
    expect(cancel).not.toHaveBeenCalled();
    await gateway.onMessageRecalled({ messageId: 'active', chatId: 'g-recall' });
    expect(cancel).toHaveBeenCalledTimes(1);
    Reflect.get(gateway, 'activeRuns').delete('group:g-recall');
    const processed: string[] = [];
    Reflect.set(gateway, 'onMessage', async (item: NormalizedMessage) => { processed.push(item.messageId); });
    await Reflect.get(gateway, 'drainQueue').call(gateway, 'g-recall');
    expect(processed).toEqual(['queued-2']);
  });

  it('persists later messages immediately while an earlier serialized request is still running', async () => {
    const { FeishuGateway } = await import('../gateway/gateway');
    const { stateDb } = await import('../state/store');
    const gateway = new FeishuGateway({ send: async () => ({ messageId: 'queue-reply' }),
      rawClient: { request: async () => ({ code: 0, data: {} }) } } as unknown as LarkChannel);
    Reflect.get(gateway, 'queues').set('g-recall', new Promise<void>(() => {}));
    Reflect.set(gateway, 'drainQueue', async () => {});
    await Reflect.get(gateway, 'enqueue').call(gateway, message('quick-followup'));
    const row = (await (await stateDb()).execute({ sql: 'SELECT status FROM queued_messages WHERE message_id = ?', args: ['quick-followup'] })).rows[0];
    expect(row?.status).toBe('pending');
  });

  it('aborts the active Mastra run and suppresses its reply after the request is recalled', async () => {
    const { FeishuGateway } = await import('../gateway/gateway');
    const sent: unknown[] = [];
    const gateway = new FeishuGateway({ send: async (_chat: string, payload: unknown) => { sent.push(payload); return { messageId: 'reply' }; },
      rawClient: { request: async () => ({ code: 0, data: {} }) } } as unknown as LarkChannel);
    let emit: ((event: unknown) => void) | undefined;
    let resolveSend: (() => void) | undefined;
    const abort = vi.fn(() => {
      emit?.({ type: 'message_update', event: { type: 'text-delta', delta: '旧回复' } });
      emit?.({ type: 'agent_end', reason: 'aborted' });
      resolveSend?.();
    });
    Reflect.set(gateway, 'sessionFor', async () => ({
      subscribe: (handler: (event: unknown) => void) => { emit = handler; return () => { emit = undefined; }; },
      sendMessage: () => new Promise<void>(resolve => { resolveSend = resolve; }), abort,
    }));
    const run = Reflect.get(gateway, 'runAgent').call(gateway, '新增任务', { userId: 'u-recall', chatId: 'g-recall', chatType: 'group', messageId: 'active-2' }, 'active-2');
    await vi.waitFor(() => expect(Reflect.get(gateway, 'activeRuns').size).toBe(1));
    await gateway.onMessageRecalled({ messageId: 'active-2', chatId: 'g-recall' });
    await run;
    await vi.waitFor(() => expect(Reflect.get(gateway, 'activeRuns').size).toBe(0));
    expect(abort).toHaveBeenCalledTimes(1);
    expect(sent).toHaveLength(0);
  });
});
