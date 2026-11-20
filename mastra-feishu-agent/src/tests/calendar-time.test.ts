import { expect, it } from 'vitest';
import { actionSummary } from '../office/actions';
import { reconcileRelativeCalendarDate } from '../office/calendar-time';
import { RequestContext } from '@mastra/core/request-context';

const requestTime = Date.parse('2026-10-26T10:55:00+08:00');
it('preserves scheduling after a title containing misleading dates and clocks', () => {
  const result = reconcileRelativeCalendarDate({ kind: 'calendar_create', summary: 'xxx',
    startTime: String(Date.parse('2027-01-15T11:20:00+08:00') / 1000),
    endTime: String(Date.parse('2027-01-15T12:20:00+08:00') / 1000),
  }, '创建一个日程，主题是明天11:20-12:20的复盘，时间为今天18:00-19:00', Date.parse('2026-11-20T10:00:00+08:00'));
  expect(result.startTime).toBe(String(Date.parse('2026-11-20T18:00:00+08:00') / 1000));
  expect(result.endTime).toBe(String(Date.parse('2026-11-20T19:00:00+08:00') / 1000));
});
const staleEvent = {
  kind: 'calendar_create' as const,
  permissionMode: 'user' as const,
  summary: '10月工作总结大会',
  startTime: '1792717200',
  endTime: '1792720800',
};

it.each(['今天11:00-12:00', '今天上午11点到12点'])('corrects the reported same-day clock drift: %s', message => {
  const action = reconcileRelativeCalendarDate({ ...staleEvent,
    startTime: String(Date.parse('2026-11-18T21:40:00+08:00') / 1000),
    endTime: String(Date.parse('2026-11-18T22:40:00+08:00') / 1000),
  }, message, Date.parse('2026-11-18T10:00:00+08:00'));
  expect(action.startTime).toBe(String(Date.parse('2026-11-18T11:00:00+08:00') / 1000));
  expect(action.endTime).toBe(String(Date.parse('2026-11-18T12:00:00+08:00') / 1000));
});

it('moves a stale model timestamp to the tomorrow requested by the Feishu user before approval', () => {
  const action = reconcileRelativeCalendarDate(staleEvent, '请为我创建一个明天上午9点到10点的会议日程，主题是10月工作总结大会', requestTime);
  expect(action).toMatchObject({ startTime: '1793062800', endTime: '1793066400' });
  expect(actionSummary({ ...action, calendarId: 'personal-calendar' })).toContain('2026-10-27 09:00');
  expect(actionSummary({ ...action, calendarId: 'personal-calendar' })).toContain('2026-10-27 10:00');
});

it('preserves an explicitly requested calendar date', () => {
  const action = reconcileRelativeCalendarDate(staleEvent, '请创建 2026 年 10 月 23 日上午 9 点的日程', requestTime);
  expect(action.startTime).toBe('1792717200');
});

it('ignores relative dates inside a meeting title while fixing the explicit clocks', () => {
  const result = reconcileRelativeCalendarDate(staleEvent, '今天11:00-12:00，主题是明天的计划', requestTime);
  expect(result.startTime).toBe(String(Date.parse('2026-10-26T11:00:00+08:00') / 1000));
});

it('does not collapse a multi-date event onto one day', () => {
  expect(reconcileRelativeCalendarDate(staleEvent, '2026-11-18 23:00至2026-11-19 01:00', requestTime)).toEqual(staleEvent);
});

it('does not shift an explicit date when tomorrow appears elsewhere in the message', () => {
  const action = reconcileRelativeCalendarDate(staleEvent, '请创建 2026 年 10 月 23 日上午 9 点的日程，标题叫明天的计划', requestTime);
  expect(action.startTime).toBe('1792717200');
});

it('uses an explicit requested date over an old timestamp', () => {
  const action = reconcileRelativeCalendarDate(staleEvent, '请创建 2026 年 10 月 27 日上午 9 点的日程', requestTime);
  expect(action.startTime).toBe('1793062800');
});

it('tells the model the current Shanghai date for each Feishu request', async () => {
  const { agent } = await import('../mastra/agents/agent');
  const requestContext = new RequestContext();
  requestContext.set('messageCreatedAt', requestTime);
  const instructions = await agent.getInstructions({ requestContext });
  expect(JSON.stringify(instructions)).toContain('2026-10-26 10:55');
});
