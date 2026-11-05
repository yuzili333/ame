import { describe, expect, it, vi } from 'vitest';
import type { NormalizedMessage } from '@larksuiteoapi/node-sdk';
import { parseSummaryRange, parseSummaryRequest } from '../gateway/summary-range';
import { isOwnerMention } from '../gateway/mention-watch';
import { readHistory, summarizeHistory } from '../gateway/chat-summary';
import type { OfficeApi } from '../office/client';

const shanghai = (value: string) => Date.parse(`${value}+08:00`);

describe('summary ranges', () => {
  it.each([
    ['本周', '2026-11-02T00:00:00'],
    ['本月', '2026-11-01T00:00:00'],
    ['本季度', '2026-10-01T00:00:00'],
    ['本年', '2026-01-01T00:00:00'],
  ])('calculates %s in Beijing time', (name, start) => {
    expect(parseSummaryRange(`总结${name}聊天`, shanghai('2026-11-03T14:00:00')).startMs).toBe(shanghai(start));
  });
  it('includes the whole end date and handles leap day', () => {
    const range = parseSummaryRange('2024年2月29日至2024年3月1日', shanghai('2024-03-02T00:00:00'));
    expect(range.startMs).toBe(shanghai('2024-02-29T00:00:00'));
    expect(range.endMs).toBe(shanghai('2024-03-02T00:00:00') - 1);
  });
  it('uses the specified end minute without extending it to the end of day', () => {
    const range = parseSummaryRange('2026年10月1日 09:15至2026年10月1日 10:30', Date.now());
    expect(range.startMs).toBe(shanghai('2026-10-01T09:15:00'));
    expect(range.endMs).toBe(shanghai('2026-10-01T10:31:00') - 1);
  });
  it('rejects reversed and invalid ranges', () => {
    expect(() => parseSummaryRange('2026年10月2日至2026年10月1日', Date.now())).toThrow();
    expect(() => parseSummaryRange('2025年2月29日至2025年3月1日', Date.now())).toThrow();
  });
  it('recognizes the named peer in a private request', () => {
    expect(parseSummaryRequest('总结本周我和张三的私聊', Date.now(), 'p2p')).toMatchObject({ kind: 'private', peerName: '张三' });
    expect(parseSummaryRequest('总结我和张三本周的聊天记录', Date.now(), 'p2p')).toMatchObject({ kind: 'private', peerName: '张三' });
    expect(parseSummaryRequest('总结本周群聊记录', Date.now(), 'group')).toMatchObject({ kind: 'group' });
  });
});

describe('mention backups', () => {
  const message = { chatType: 'group', senderId: 'ou_sender', mentionAll: false, mentions: [{ key: '@_user_1', openId: 'ou_owner' }] } as NormalizedMessage;
  it('matches only explicit owner mentions', () => {
    expect(isOwnerMention(message, 'ou_owner')).toBe(true);
    expect(isOwnerMention({ ...message, mentionAll: true }, 'ou_owner')).toBe(false);
    expect(isOwnerMention({ ...message, senderId: 'ou_owner' }, 'ou_owner')).toBe(false);
    expect(isOwnerMention({ ...message, mentions: [] }, 'ou_owner')).toBe(false);
  });
});

describe('history pagination', () => {
  it('reads all pages and filters deleted and out-of-range messages', async () => {
    const calls: unknown[] = [];
    const office = { call: async (request: unknown) => {
      calls.push(request);
      return calls.length === 1
        ? { items: [{ message_id: 'a', create_time: '1000' }, { message_id: 'x', create_time: '1001', deleted: true }, { message_id: 'bot', create_time: '1001', sender: { sender_type: 'app' } }], has_more: true, page_token: 'next' }
        : { items: [{ message_id: 'b', create_time: '1002' }], has_more: false };
    } } as unknown as OfficeApi;
    const result = await readHistory(office, 'oc_chat', { startMs: 1000, endMs: 1002, label: 'test' });
    expect(result.messages.map(item => item.message_id)).toEqual(['a', 'b']);
    expect(result.truncated).toBe(false);
    expect(calls).toHaveLength(2);
  });
});

describe('summary presentation', () => {
  it('accepts streaming data frames from the intranet model proxy', async () => {
    const stream = [
      'data: {"choices":[{"delta":{"content":"### 📋 工作事项\\n"}}]}\n\n',
      'data: {"choices":[{"delta":{"content":"- 09:00 确认排期"}}]}\n\n',
      'data: [DONE]\n\n',
    ].join('');
    const fetchMock = vi.fn(async (_url: string, _options: RequestInit) => new Response(stream, { status: 200, headers: { 'content-type': 'text/event-stream' } }));
    vi.stubGlobal('fetch', fetchMock);
    try {
      const result = await summarizeHistory([{ message_id: 'om_1', create_time: String(shanghai('2026-11-03T09:00:00')), body: { content: '{"text":"确认排期"}' } }],
        { startMs: 0, endMs: Date.now(), label: '本周' }, false);
      expect(result).toContain('### 📋 工作事项\n- 09:00 确认排期');
      const request = JSON.parse(String(fetchMock.mock.calls[0]![1].body)) as { stream?: boolean };
      expect(request.stream).toBe(true);
    } finally { vi.unstubAllGlobals(); }
  });

  it('omits empty technical categories and a standalone mention section', async () => {
    const fetchMock = vi.fn(async (_url: string, _options: RequestInit) => new Response(JSON.stringify({ choices: [{ message: { content: '### 🛠️ 技术问题\n输入消息中无相关技术问题、Bug 修复或问题解答内容。\n\n### 成员@消息\n- @张三提醒处理排期\n\n### 📋 工作事项\n- 09:00 李四 @张三确认排期。' } }] }), { status: 200, headers: { 'content-type': 'application/json' } }));
    vi.stubGlobal('fetch', fetchMock);
    try {
      const result = await summarizeHistory([{ message_id: 'om_1', create_time: String(shanghai('2026-11-03T09:00:00')), sender: { id: 'ou_1', sender_name: '李四' }, body: { content: '{"text":"@张三确认排期"}' } }],
        { startMs: 0, endMs: Date.now(), label: '本周' }, false);
      expect(result).toContain('### 📋 工作事项');
      expect(result).toContain('@张三确认排期');
      expect(result).not.toContain('### 成员@消息');
      expect(result).not.toContain('无相关技术问题');
      expect(result).not.toContain('### 🛠️ 技术问题');
      const request = JSON.parse(String(fetchMock.mock.calls[0]![1].body)) as { messages: Array<{ content: string }> };
      expect(request.messages[0].content).toContain('没有内容的分类直接省略');
    } finally { vi.unstubAllGlobals(); }
  });
});
