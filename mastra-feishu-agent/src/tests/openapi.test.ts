import { expect, it } from 'vitest';
import { discoverApis, prepareApiRequest, apiOperation, type OpenApiInput } from '../office/openapi';
import { managedOperations } from '../office/managed-operations';

const invite: OpenApiInput = { operationId: 'calendar.v4.calendar.event.attendee.create', permissionMode: 'user', scopes: [], summary: '邀请同事参加日程', arguments: {
  path: { calendar_id: 'calendar@group.calendar.feishu.cn', event_id: 'event_0' }, params: { user_id_type: 'open_id' }, data: { attendees: [{ type: 'user', user_id: 'ou_colleague' }], need_notification: true },
} };
it('discovers SDK APIs and exposes their actual request schema', () => {
  expect(discoverApis('').total).toBeGreaterThan(1000);
  expect(discoverApis('搜索用户').operations).toContainEqual(expect.objectContaining({ id: 'contact.user.search' }));
  expect(apiOperation(invite.operationId).input).toMatchObject({ type: 'object', required: ['path'] });
  expect(prepareApiRequest(invite)).toMatchObject({ method: 'POST', url: '/open-apis/calendar/v4/calendars/calendar%40group.calendar.feishu.cn/events/event_0/attendees', params: { user_id_type: 'open_id' }, data: invite.arguments.data });
});
it('rejects unknown APIs, invalid IDs, injected headers, credentials, and wrong parameter types', () => {
  expect(() => prepareApiRequest({ ...invite, operationId: 'https://evil.test' })).toThrow();
  expect(() => apiOperation('drive.v1.file.download')).toThrow();
  expect(() => apiOperation('acs.v1.access_record.access_photo.get')).toThrow();
  expect(() => prepareApiRequest({ ...invite, arguments: { ...invite.arguments, path: { calendar_id: '..', event_id: 'e' } } })).toThrow();
  expect(() => prepareApiRequest({ ...invite, arguments: { ...invite.arguments, headers: { Authorization: 'secret' } } as never })).toThrow();
  expect(() => prepareApiRequest({ ...invite, arguments: { ...invite.arguments, data: { access_token: 'secret' } } })).toThrow();
  expect(() => prepareApiRequest({ ...invite, arguments: { ...invite.arguments, params: { user_id_type: 'invented' } } })).toThrow();
  expect(() => prepareApiRequest({ ...invite, arguments: {} })).toThrow();
});
it('does not let generic calendar writes bypass time normalization', () => {
  expect(() => prepareApiRequest({ ...invite, operationId: 'calendar.v4.calendar.event.create' })).toThrow('专用工具');
});
it('keeps other calendar APIs available without adding dedicated tool functions', () => {
  expect(prepareApiRequest({ ...invite, operationId: 'calendar.v4.calendar.acl.create', arguments: {
    path: { calendar_id: 'cal' }, params: { user_id_type: 'open_id' }, data: { role: 'reader', scope: { type: 'user', user_id: 'ou1' } },
  } })).toMatchObject({ method: 'POST', url: '/open-apis/calendar/v4/calendars/cal/acls' });
});
it('searches colleagues with requester identity and preserves pagination', () => {
  const input: OpenApiInput = { operationId: 'contact.user.search', permissionMode: 'user', summary: '查找张三', scopes: [], arguments: { params: { query: '张三', page_token: 'next' } } };
  expect(prepareApiRequest(input)).toMatchObject({ method: 'GET', url: '/open-apis/search/v1/user', params: { query: '张三', page_token: 'next' } });
  expect(() => prepareApiRequest({ ...input, permissionMode: 'app' })).toThrow('用户身份');
});

it('directs covered raw writes to the same business handlers instead of bypassing validation', async () => {
  const { officeBusinessTools } = await import('../mastra/tools/office-tools');
  const ids = new Set(Object.values(officeBusinessTools).map(tool => tool.id));
  for (const [operationId, businessId] of Object.entries(managedOperations)) {
    expect(ids.has(businessId)).toBe(true);
    expect(discoverApis('', operationId)).toMatchObject({ managedOperationId: businessId });
    expect(() => prepareApiRequest({ ...invite, operationId })).toThrow(businessId);
  }
});
