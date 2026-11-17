import { expect, it } from 'vitest';
import { reconcileTaskTime } from '../office/task-time';

const ms = (local: string) => String(Date.parse(`${local}+08:00`));
const receivedAt = Number(ms('2026-11-17T00:05:00'));
const action = { kind: 'task_create' as const, summary: 'Work', dueTimestamp: ms('2026-10-26T17:00:00') };

it.each([
  ['今天下午16:00-17:00', '2026-11-17'],
  ['明天16:00至17:00', '2026-11-18'],
  ['后天16：00～17：00', '2026-11-19'],
  ['2026年12月1日16:00—17:00', '2026-12-01'],
  ['2026-12-01 16:00-17:00', '2026-12-01'],
])('anchors %s in Shanghai and extracts the complete time range', (message, date) => {
  expect(reconcileTaskTime(action, message, receivedAt)).toMatchObject({ startTimestamp: ms(`${date}T16:00:00`), dueTimestamp: ms(`${date}T17:00:00`) });
});

it('uses the stated hours instead of a stale model clock, including afternoon notation', () => {
  expect(reconcileTaskTime(action, '今天下午4:30-5:45', receivedAt)).toMatchObject({ startTimestamp: ms('2026-11-17T16:30:00'), dueTimestamp: ms('2026-11-17T17:45:00') });
});

it('corrects a deadline-only date without inventing a start', () => {
  const result = reconcileTaskTime(action, '今天17点前完成', receivedAt);
  expect(result.dueTimestamp).toBe(ms('2026-11-17T17:00:00'));
  expect(result).not.toHaveProperty('startTimestamp');
});

it('preserves an overnight duration when only the date needs reconciling', () => {
  const result = reconcileTaskTime({ ...action, startTimestamp: ms('2026-10-26T23:00:00'), dueTimestamp: ms('2026-10-27T01:00:00') }, '今天开始，持续两小时', receivedAt);
  expect(result).toMatchObject({ startTimestamp: ms('2026-11-17T23:00:00'), dueTimestamp: ms('2026-11-18T01:00:00') });
});

it('ignores dates in the title and preserves historical dates in user content', () => {
  const description = '复盘2026-10-26的结果';
  const result = reconcileTaskTime({ ...action, description }, '今天16:00-17:00，主题是编写明天的计划和复盘2026-10-26的结果', receivedAt);
  expect(result).toMatchObject({ description, startTimestamp: ms('2026-11-17T16:00:00'), dueTimestamp: ms('2026-11-17T17:00:00') });
});

it.each(['今天还是明天16:00-17:00', '11月20日16:00-17:00', '今天11/20的任务', '今天16:00-17:00，明天18:00-19:00', '今天16:00-17:00和18:00-19:00'])('does not reinterpret unsupported or multiple date ranges: %s', message => {
  expect(reconcileTaskTime(action, message, receivedAt)).toEqual(action);
});

it.each(['今天17:00-16:00', '今天25:00-26:00', '2026-02-30 16:00-17:00'])('rejects invalid explicit times before approval: %s', message => {
  expect(() => reconcileTaskTime(action, message, receivedAt)).toThrow();
});

it('changes only the supplied endpoint of a task update', () => {
  expect(reconcileTaskTime({ kind: 'task_update', taskId: 'task1', startTimestamp: ms('2026-10-26T16:00:00') }, '改为今天16点开始', receivedAt)).toEqual({ kind: 'task_update', taskId: 'task1', startTimestamp: ms('2026-11-17T16:00:00') });
});

it('preserves scheduling notes that are not a generated timestamp-only line', () => {
  const description = '时间窗口：请先和参与人确认，再预留一小时';
  expect(reconcileTaskTime({ ...action, description }, '今天16:00-17:00', receivedAt).description).toBe(description);
});
