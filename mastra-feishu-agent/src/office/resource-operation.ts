import { z } from 'zod';
import { officeAction, type OfficeAction } from './actions';
import type { OfficeApi } from './client';
import { assertTaskTimeRange } from './task-time';

const token = z.string().min(1).max(256).regex(/^[\w@.-]+$/);
const timestamp = z.string().regex(/^\d{10}$/);
const identity = z.enum(['user', 'app']).optional();
const page = { pageToken: z.string().min(1).max(4096).optional(), pageSize: z.number().int().min(1).max(100).optional() };
export const resourceRead = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('calendar_list'), permissionMode: identity, ...page }),
  z.object({ kind: z.literal('calendar_read'), calendarId: token.optional(), permissionMode: identity }),
  z.object({ kind: z.literal('calendar_events'), calendarId: token.optional(), startTime: timestamp, endTime: timestamp, pageToken: page.pageToken, permissionMode: identity }),
  z.object({ kind: z.literal('calendar_event_read'), calendarId: token, eventId: token, permissionMode: identity }),
  z.object({ kind: z.literal('task_list'), completed: z.boolean().optional(), ...page }),
  z.object({ kind: z.literal('task_read'), taskId: token }),
]);
export type ResourceRead = z.infer<typeof resourceRead>;
export type ResourceWrite = Extract<OfficeAction, { kind: `calendar_${string}` | `task_${string}` }>;
export type ResourceIntent = ResourceRead | ResourceWrite;
export function isResourceWrite(action: { kind: string }): action is ResourceWrite {
  return ['calendar_create', 'calendar_update', 'calendar_delete', 'calendar_container_create', 'calendar_container_update', 'calendar_container_delete', 'task_create', 'task_update', 'task_delete'].includes(action.kind);
}
export function parseResourceIntent(input: unknown): ResourceIntent {
  const read = resourceRead.safeParse(input);
  if (read.success) return read.data;
  const action = officeAction.parse(input);
  if (!isResourceWrite(action)) throw new Error('不是日历或任务操作');
  return action;
}
export const resourceFamily = (intent: ResourceIntent) => intent.kind.startsWith('task_') ? 'task' as const : 'calendar' as const;
export const resourceMode = (intent: ResourceIntent): 'app' | 'user' => 'permissionMode' in intent && intent.permissionMode === 'app' ? 'app' : 'user';
export function requiredResourceScopes(intent: ResourceIntent): string[] {
  if (resourceFamily(intent) === 'task') return [isResourceWrite(intent) ? 'task:task:write' : 'task:task:read', 'offline_access'];
  return isResourceWrite(intent) ? ['calendar:calendar', 'calendar:calendar:readonly', 'offline_access'] : ['calendar:calendar:readonly', 'offline_access'];
}
export class ResourceAccessError extends Error {}
const encoded = encodeURIComponent;
export async function calendarInfo(api: OfficeApi, calendarId: string, userAccessToken?: string) {
  const data = await api.call({ method: 'GET', url: `/open-apis/calendar/v4/calendars/${encoded(calendarId)}`, userAccessToken });
  // Current API returns the entity directly; older private deployments wrap it.
  return (data.calendar ?? data) as { calendar_id?: string; type?: string; role?: string; summary?: string };
}

export async function prepareResourceWrite(api: OfficeApi, input: ResourceWrite, userAccessToken?: string): Promise<ResourceWrite> {
  let action = { ...input, permissionMode: resourceMode(input) } as ResourceWrite;
  if (action.kind.startsWith('calendar_')) {
    if (action.kind === 'calendar_container_create') return action;
    if (action.kind === 'calendar_create' && !action.calendarId) {
      if (!userAccessToken) throw new Error('应用身份需要明确的共享日历 ID');
      action = { ...action, calendarId: await api.primaryCalendar(userAccessToken) };
    }
    if ('calendarId' in action && action.calendarId) {
      const info = await calendarInfo(api, action.calendarId, userAccessToken);
      if (action.kind === 'calendar_container_delete' && info.type !== 'shared') throw new Error('只能删除共享日历，不能删除个人主日历或其他类型日历');
      const roles = action.kind.startsWith('calendar_container_') ? ['owner'] : ['owner', 'writer'];
      if (!roles.includes(info.role ?? '')) throw new ResourceAccessError('当前身份没有目标日历的操作权限，请联系日历所有者授予相应权限');
    }
    if (action.kind === 'calendar_create' && Number(action.endTime) <= Number(action.startTime)) throw new Error('日程结束时间必须晚于开始时间');
    if (action.kind === 'calendar_update') {
      if (![action.summary, action.description, action.startTime, action.endTime].some(value => value !== undefined)) throw new Error('请提供要修改的日程字段');
      if (action.startTime || action.endTime) {
        const data = await api.call({ method: 'GET', url: `/open-apis/calendar/v4/calendars/${encoded(action.calendarId)}/events/${encoded(action.eventId)}`, userAccessToken });
        const event = data.event as { start_time?: { timestamp?: string }; end_time?: { timestamp?: string } } | undefined;
        const start = action.startTime ?? event?.start_time?.timestamp;
        const end = action.endTime ?? event?.end_time?.timestamp;
        if (!start || !end || Number(end) <= Number(start)) throw new Error('请提供有效的完整开始和结束时间；结束时间必须晚于开始时间');
      }
    }
    if (action.kind === 'calendar_container_update' && ![action.summary, action.description, action.permissions].some(value => value !== undefined)) throw new Error('请提供要修改的日历字段');
  } else if (action.kind === 'task_create') {
    assertTaskTimeRange(action);
  } else if ('taskId' in action) {
    const data = await api.readTask(action.taskId, userAccessToken);
    if (action.kind === 'task_update') {
      if (![action.summary, action.description, action.startTimestamp, action.dueTimestamp, action.completedAt].some(value => value !== undefined)) throw new Error('请提供要修改的任务字段');
      if (action.startTimestamp || action.dueTimestamp) {
        const task = data.task as { start?: { timestamp?: string }; due?: { timestamp?: string } } | undefined;
        assertTaskTimeRange({ startTimestamp: action.startTimestamp ?? task?.start?.timestamp, dueTimestamp: action.dueTimestamp ?? task?.due?.timestamp });
      }
    }
  }
  return action;
}

export async function readResource(api: OfficeApi, intent: ResourceRead, userAccessToken?: string): Promise<Record<string, unknown>> {
  const pagination = 'pageSize' in intent || 'pageToken' in intent
    ? { page_size: 'pageSize' in intent ? intent.pageSize ?? 50 : 50, ...('pageToken' in intent && intent.pageToken ? { page_token: intent.pageToken } : {}) } : {};
  switch (intent.kind) {
    case 'task_list':
      return api.call({ method: 'GET', url: '/open-apis/task/v2/tasks', userAccessToken, params: { type: 'my_tasks', ...pagination, ...(intent.completed !== undefined ? { completed: intent.completed } : {}) } });
    case 'task_read': return api.readTask(intent.taskId, userAccessToken);
    case 'calendar_list': return api.call({ method: 'GET', url: '/open-apis/calendar/v4/calendars', userAccessToken, params: pagination });
    case 'calendar_event_read': return api.call({ method: 'GET', url: `/open-apis/calendar/v4/calendars/${encoded(intent.calendarId)}/events/${encoded(intent.eventId)}`, userAccessToken });
    case 'calendar_events':
    case 'calendar_read': {
      if (!intent.calendarId && !userAccessToken) throw new Error('应用身份需要明确的共享日历 ID');
      const calendarId = intent.calendarId ?? await api.primaryCalendar(userAccessToken!, false);
      if (intent.kind === 'calendar_read') return calendarInfo(api, calendarId, userAccessToken);
      if (Number(intent.endTime) <= Number(intent.startTime)) throw new Error('查询结束时间必须晚于开始时间');
      return api.readCalendar(calendarId, intent.startTime, intent.endTime, userAccessToken, intent.pageToken);
    }
  }
}

export function resourceReadSummary(intent: ResourceRead): string {
  switch (intent.kind) {
    case 'calendar_list': return '查询日历列表';
    case 'calendar_read': return `查询日历 ${intent.calendarId ?? '个人主日历'}`;
    case 'calendar_events': return `查询日历 ${intent.calendarId ?? '个人主日历'} 的日程`;
    case 'calendar_event_read': return `查询日历 ${intent.calendarId} 的日程 ${intent.eventId}`;
    case 'task_list': return '查询我负责的任务列表';
    case 'task_read': return `查询任务 ${intent.taskId}`;
  }
}
