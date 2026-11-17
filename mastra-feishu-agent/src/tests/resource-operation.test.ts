import { expect, it } from 'vitest';
import { OfficeApi, type ApiRequest } from '../office/client';
import { officeAction } from '../office/actions';
import { prepareResourceWrite, readResource } from '../office/resource-operation';

function fixture() {
  const requests: ApiRequest[] = [];
  const api = new OfficeApi({ request: async request => { requests.push(request); return { code: 0, data: { task: { guid: 'task1' } } }; }, send: async () => undefined });
  return { api, requests };
}

it.each([
  [{ kind: 'calendar_delete', calendarId: 'cal', eventId: 'evt', permissionMode: 'user' }, '/open-apis/calendar/v4/calendars/cal/events/evt'],
  [{ kind: 'calendar_container_delete', calendarId: 'cal', permissionMode: 'user' }, '/open-apis/calendar/v4/calendars/cal'],
  [{ kind: 'task_delete', taskId: 'task1', permissionMode: 'user' }, '/open-apis/task/v2/tasks/task1'],
])('deletes the explicit resource using the user token: %j', async (input, url) => {
  const { api, requests } = fixture();
  await api.execute(officeAction.parse(input), 'key', 'user-token');
  expect(requests).toMatchObject([{ method: 'DELETE', url, userAccessToken: 'user-token' }]);
});

it('creates and edits the requester task using user identity and completed_at', async () => {
  const { api, requests } = fixture();
  await api.execute(officeAction.parse({ kind: 'task_create', summary: 'Mine', permissionMode: 'user' }), 'unique-key', 'user-token');
  await api.execute(officeAction.parse({ kind: 'task_update', taskId: 'task1', completedAt: '0', permissionMode: 'user' }), 'key', 'user-token');
  expect(requests).toMatchObject([
    { method: 'POST', userAccessToken: 'user-token', data: { summary: 'Mine', client_token: 'unique-key' } },
    { method: 'PATCH', userAccessToken: 'user-token', data: { task: { completed_at: '0' }, update_fields: ['completed_at'] } },
  ]);
});

it('never falls back to app identity for user writes', async () => {
  const { api, requests } = fixture();
  await expect(api.execute(officeAction.parse({ kind: 'task_create', summary: 'Mine', permissionMode: 'user' }))).rejects.toThrow('用户');
  expect(requests).toEqual([]);
});

it('creates and edits a calendar container independently from its events', async () => {
  const { api, requests } = fixture();
  await api.execute(officeAction.parse({ kind: 'calendar_container_create', summary: 'Team', permissionMode: 'user' }), 'key', 'user-token');
  await api.execute(officeAction.parse({ kind: 'calendar_container_update', calendarId: 'cal', description: 'New', permissionMode: 'user' }), 'key', 'user-token');
  expect(requests).toMatchObject([
    { method: 'POST', url: '/open-apis/calendar/v4/calendars', data: { summary: 'Team' }, userAccessToken: 'user-token' },
    { method: 'PATCH', url: '/open-apis/calendar/v4/calendars/cal', data: { description: 'New' }, userAccessToken: 'user-token' },
  ]);
});

it('reads primary calendar events with read-only access and preserves pagination', async () => {
  const requests: ApiRequest[] = [];
  const api = new OfficeApi({ request: async request => {
    requests.push(request);
    return request.url.endsWith('/primary')
      ? { code: 0, data: { calendars: [{ calendar: { calendar_id: 'primary', type: 'primary', role: 'reader' } }] } }
      : { code: 0, data: { items: [{ event_id: 'evt' }], has_more: true, page_token: 'next-page' } };
  }, send: async () => undefined });
  const result = await readResource(api, { kind: 'calendar_events', startTime: '1760000000', endTime: '1760003600', pageToken: 'first-page' }, 'read-token');
  expect(result).toMatchObject({ has_more: true, page_token: 'next-page', items: [{ event_id: 'evt' }] });
  expect(requests).toMatchObject([{ userAccessToken: 'read-token' }, { userAccessToken: 'read-token', url: '/open-apis/calendar/v4/calendars/primary/events', params: { page_token: 'first-page' } }]);
});

it('uses direct calendar get responses when resolving a personal event write', async () => {
  const api = new OfficeApi({ request: async request => request.url.endsWith('/primary')
    ? { code: 0, data: { calendars: [{ calendar: { calendar_id: 'mine', type: 'primary', role: 'owner' } }] } }
    : { code: 0, data: { calendar_id: 'mine', type: 'primary', role: 'owner' } }, send: async () => undefined });
  expect(await prepareResourceWrite(api, { kind: 'calendar_create', summary: 'New event', startTime: '1760000000', endTime: '1760003600' }, 'user-token')).toMatchObject({ calendarId: 'mine', permissionMode: 'user' });
});

it('validates a partially edited event range against the actual unchanged endpoint', async () => {
  const api = new OfficeApi({ request: async request => ({ code: 0, data: request.url.includes('/events/')
    ? { event: { start_time: { timestamp: '1760000000' }, end_time: { timestamp: '1760003600' } } }
    : { calendar_id: 'cal', role: 'writer' } }), send: async () => undefined });
  await expect(prepareResourceWrite(api, { kind: 'calendar_update', calendarId: 'cal', eventId: 'evt', startTime: '1760007200' }, 'token')).rejects.toThrow('结束时间');
});

it('normalizes bare HTTP permission denials so an unauthorized selected calendar can trigger consent', async () => {
  const api = new OfficeApi({ request: async () => { throw Object.assign(new Error('Forbidden'), { response: { status: 403 } }); }, send: async () => undefined });
  await expect(api.call({ method: 'GET', url: '/open-apis/calendar/v4/calendars/cal' })).rejects.toThrow('Feishu API 403');
});

it('sends a task start time on updates instead of dropping it from the schema', async () => {
  const { api, requests } = fixture();
  const startTimestamp = String(Date.parse('2026-11-17T16:00:00+08:00'));
  await api.execute(officeAction.parse({ kind: 'task_update', taskId: 'task1', startTimestamp, permissionMode: 'user' }), 'key', 'user-token');
  expect(requests).toMatchObject([{ method: 'PATCH', data: { task: { start: { timestamp: startTimestamp, is_all_day: false } }, update_fields: ['start'] } }]);
});

it('checks partial task time edits against the unchanged remote endpoint', async () => {
  const api = new OfficeApi({ request: async () => ({ code: 0, data: { task: {
    start: { timestamp: String(Date.parse('2026-11-17T16:00:00+08:00')) },
    due: { timestamp: String(Date.parse('2026-11-17T17:00:00+08:00')) },
  } } }), send: async () => undefined });
  await expect(prepareResourceWrite(api, { kind: 'task_update', taskId: 'task1', startTimestamp: String(Date.parse('2026-11-17T18:00:00+08:00')) }, 'token')).rejects.toThrow('截止时间');
  await expect(prepareResourceWrite(api, { kind: 'task_update', taskId: 'task1', dueTimestamp: String(Date.parse('2026-11-17T15:00:00+08:00')) }, 'token')).rejects.toThrow('截止时间');
});
