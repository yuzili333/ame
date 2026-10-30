import { expect, it } from 'vitest';
import { actionSummary } from '../office/actions';
import { reconcileRelativeCalendarDate } from '../office/calendar-time';
import { RequestContext } from '@mastra/core/request-context';

const requestTime = Date.parse('2026-10-26T10:55:00+08:00');
const staleEvent = {
  kind: 'calendar_create' as const,
  permissionMode: 'user' as const,
  summary: '10月工作总结大会',
  startTime: '1792717200',
  endTime: '1792720800',
};

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
