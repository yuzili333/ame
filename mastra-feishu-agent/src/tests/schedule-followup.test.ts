import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';

process.env.AGENT_DATA_DIR = mkdtempSync(join(tmpdir(), 'schedule-followup-'));
const caller = { userId: 'u1', chatId: 'c1', chatType: 'p2p' as const, messageId: 'original-schedule' };
const time = Date.parse('2026-11-20T23:58:00+08:00');

it('carries the original schedule through a type correction across midnight, isolated by user and chat', async () => {
  const { resolveScheduleSource } = await import('../office/schedule-followup');
  const originalMessage = '创建一个任务，主题是xxx，时间为今天18:00-19:00';
  await resolveScheduleSource(caller, originalMessage, time);
  const result = await resolveScheduleSource({ ...caller, messageId: 'followup' }, '改为创建日程', time + 300_000);
  expect(result).toMatchObject({ messageCreatedAt: time });
  expect(result.originalMessage).toContain(originalMessage);
  expect(result.originalMessage).toContain('改为创建日程');
  for (const isolated of [{ ...caller, userId: 'u2' }, { ...caller, chatId: 'c2' }]) {
    expect(await resolveScheduleSource(isolated, '改为创建日程', time + 300_000)).toMatchObject({ needsTime: true });
  }
  expect(await resolveScheduleSource(caller, '改为创建日程', time + 31 * 60_000)).toMatchObject({ needsTime: true });
});

it('uses an explicitly supplied new time and ignores recalled sources', async () => {
  const { resolveScheduleSource } = await import('../office/schedule-followup');
  const { stateDb } = await import('../state/store');
  const source = { ...caller, chatId: 'new-time', messageId: 'new-source' };
  await resolveScheduleSource(source, '创建任务，时间为今天18:00-19:00', time);
  const newMessage = '改为日程，时间为明天10:00-11:00';
  expect(await resolveScheduleSource({ ...source, messageId: 'new-time-message' }, newMessage, time + 1000))
    .toMatchObject({ originalMessage: newMessage, messageCreatedAt: time + 1000 });
  await (await stateDb()).execute({ sql: 'INSERT INTO recalled_messages(message_id,chat_id,recalled_at) VALUES (?,?,?)', args: ['new-time-message', source.chatId, time + 2000] });
  expect(await resolveScheduleSource(source, '改为创建日程', time + 3000)).toMatchObject({ needsTime: true });
});
