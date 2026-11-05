import { afterAll, beforeAll, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createClient } from '@libsql/client';
import type { LarkChannel } from '@larksuiteoapi/node-sdk';

const dataDir = mkdtempSync(join(tmpdir(), 'feishu-gateway-trace-test-'));
let FeishuGateway: typeof import('../gateway/gateway').FeishuGateway;
let backupMention: typeof import('../gateway/mention-watch').backupMention;

beforeAll(async () => {
  vi.stubEnv('AGENT_DATA_DIR', dataDir);
  vi.resetModules();
  ({ FeishuGateway } = await import('../gateway/gateway'));
  ({ backupMention } = await import('../gateway/mention-watch'));
});
afterAll(() => { vi.unstubAllEnvs(); rmSync(dataDir, { recursive: true, force: true }); });

it('stores a directly handled chat summary as a Studio-visible trace', async () => {
  const sent: unknown[] = [];
  const gateway = new FeishuGateway({ send: async (_chat: string, payload: unknown) => { sent.push(payload); return { messageId: 'reply' }; } } as unknown as LarkChannel);
  vi.spyOn(gateway.office, 'call').mockResolvedValue({ items: [], has_more: false });
  await gateway.onMessage({
    messageId: 'summary-trace-message', senderId: 'ou_requester', chatId: 'oc_summary', chatType: 'group',
    content: '总结本周群聊记录', mentionedBot: true, rawContentType: 'text', resources: [], mentions: [], mentionAll: false, createTime: Date.now(),
  });
  expect(sent).toHaveLength(1);
  const db = createClient({ url: `file:${join(dataDir, 'traces.db')}` });
  const rows = await db.execute("SELECT name, spanType, entityType, entityId, entityName, input, output FROM mastra_ai_spans WHERE parentSpanId IS NULL ORDER BY startedAt DESC LIMIT 1");
  expect(rows.rows[0]?.name).toContain('gateway chat message');
  expect(rows.rows[0]?.entityType).toBe('agent');
  expect(rows.rows[0]?.entityId).toBe('office-agent');
  expect(rows.rows[0]?.entityName).toBe('飞书办公智能体');
  expect(Buffer.from(rows.rows[0]?.input as ArrayBuffer).toString()).toContain('总结本周群聊记录');
  expect(Buffer.from(rows.rows[0]?.output as ArrayBuffer).toString()).toContain('没有可读取的聊天消息');
  db.close();
});

it('stores a delivered owner mention backup under the office agent', async () => {
  const sent: unknown[] = [];
  const channel = { rawClient: { request: async (request: unknown) => { sent.push(request); return { code: 0 }; } } } as unknown as LarkChannel;
  const office = { call: async (request: { url: string }) => {
    if (request.url.endsWith('/members')) return { items: [{ member_id: 'ou_owner', name: '持有者' }, { member_id: 'ou_sender', name: '张三' }], has_more: false };
    if (request.url.endsWith('/oc_mention')) return { name: '项目群' };
    return { items: [{ message_position: '42' }] };
  } } as unknown as import('../office/client').OfficeApi;
  await backupMention({ messageId: 'mention-trace-message', senderId: 'ou_sender', chatId: 'oc_mention', chatType: 'group',
    content: '@持有者 请查看', mentionedBot: false, rawContentType: 'text', resources: [],
    mentions: [{ key: '@_user_1', openId: 'ou_owner', name: '持有者' }], mentionAll: false, createTime: Date.now() }, 'ou_owner', office, channel);
  expect(sent).toHaveLength(1);
  const db = createClient({ url: `file:${join(dataDir, 'traces.db')}` });
  const rows = await db.execute("SELECT name, entityType, entityId, entityName, input, output FROM mastra_ai_spans WHERE parentSpanId IS NULL AND name = 'Feishu owner mention backup' ORDER BY startedAt DESC LIMIT 1");
  expect(rows.rows[0]?.entityType).toBe('agent');
  expect(rows.rows[0]?.entityId).toBe('office-agent');
  expect(rows.rows[0]?.entityName).toBe('飞书办公智能体');
  expect(Buffer.from(rows.rows[0]?.input as ArrayBuffer).toString()).toContain('mention-trace-message');
  expect(Buffer.from(rows.rows[0]?.output as ArrayBuffer).toString()).toContain('sent');
  db.close();
});
