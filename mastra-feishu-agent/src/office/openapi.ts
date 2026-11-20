import { z } from 'zod';
import catalogData from './feishu-api-catalog.cjs';
import type { ApiRequest, OfficeApi } from './client';
import { managedOperation } from './managed-operations';

type Operation = { id: string; description: string; documentation: string; method: ApiRequest['method']; path: string; input: Record<string, unknown> };
const operations = new Map((catalogData.operations as unknown as Operation[]).map(op => [op.id, op]));
// Legacy search is documented by Feishu but absent from the current SDK.
operations.set('contact.user.search', { id: 'contact.user.search', description: '按姓名关键词搜索用户（通讯录、同事、人员），使用请求者身份；返回 open_id、姓名、部门，重名时需用户选择。', method: 'GET', path: '/open-apis/search/v1/user',
  documentation: 'https://open.feishu.cn/document/server-docs/contact-v3/user/search-users',
  input: { type: 'object', properties: { params: { type: 'object', properties: { query: { type: 'string', minLength: 1 }, page_size: { type: 'integer', minimum: 1, maximum: 100 }, page_token: { type: 'string' } }, required: ['query'], additionalProperties: false } }, required: ['params'], additionalProperties: false } });

export const openApiFields = {
  operationId: z.string().min(1).max(300),
  arguments: z.object({ path: z.record(z.string(), z.string()).optional(), params: z.record(z.string(), z.unknown()).optional(), data: z.record(z.string(), z.unknown()).optional() }).strict(),
  permissionMode: z.enum(['app', 'user']),
  scopes: z.array(z.union([z.literal('offline_access'), z.string().regex(/^[a-zA-Z][\w.-]*:[\w:.-]+$/)])).max(50).default([]),
  summary: z.string().trim().min(1).max(1200),
};
export const openApiInput = z.object(openApiFields);
export type OpenApiInput = z.infer<typeof openApiInput>;
const credentialPath = /\/(?:auth|authen)\/|\/(?:access_token|refresh_token|tenant_access_token|app_access_token|token|app_secret|client_secret|secret|ticket)(?:\/|$)/;

export function apiOperation(id: string): Operation {
  const op = operations.get(id);
  if (!op || credentialPath.test(op.path)) throw new Error('该接口不在可调用业务 API 目录中，请先检索目录');
  return op;
}
export function discoverApis(query: string, operationId?: string, offset = 0) {
  if (operationId) {
    const op = apiOperation(operationId);
    return { sdkVersion: catalogData.sdkVersion, operation: op, managedOperationId: managedOperation(op.id, op.method, op.path), scopes: knownApiScopes(operationId), note: '有 managedOperationId 时请在统一入口使用该业务操作及其参数 schema，复用已有校验。目录不代表已授权。' };
  }
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  const matches = [...operations.values()].filter(op => !credentialPath.test(op.path) && terms.every(term => `${op.id} ${op.path} ${op.description}`.toLowerCase().includes(term)));
  return { total: matches.length, hasMore: offset + 15 < matches.length, nextOffset: offset + 15,
    operations: matches.slice(offset, offset + 15).map(({ id, method, path, description }) => ({ id, method, path, managedOperationId: managedOperation(id, method, path), description: description.slice(0, 350) })) };
}
export function knownApiScopes(id: string): string[] {
  const op = apiOperation(id);
  if (id === 'contact.user.search') return ['contact:user:search'];
  if (op.path.startsWith('/open-apis/calendar/')) return [op.method === 'GET' ? 'calendar:calendar:readonly' : 'calendar:calendar'];
  if (op.path.startsWith('/open-apis/task/')) return [op.method === 'GET' ? 'task:task:read' : 'task:task:write'];
  return [];
}
export function openApiScopes(input: OpenApiInput) { return [...new Set([...knownApiScopes(input.operationId), ...input.scopes, 'offline_access'])]; }
export function prepareApiRequest(input: OpenApiInput): ApiRequest {
  const op = apiOperation(input.operationId);
  if (input.operationId === 'contact.user.search' && input.permissionMode !== 'user') throw new Error('按姓名搜索用户只支持用户身份，请使用 permissionMode=user');
  const managed = managedOperation(op.id, op.method, op.path);
  if (managed) throw new Error(`请在统一入口使用业务 operationId=${managed} 并先查询其 schema；原专用工具的时间校正、数据预检和审批由该操作复用，不能通过原始 API 绕过`);
  const raw = JSON.stringify(input.arguments);
  if (raw.length > 8_000) throw new Error('请求参数过大，无法在审批卡完整展示，请缩小操作范围');
  if (/"(?:authorization|access_token|refresh_token|client_secret|app_secret|tenant_access_token|app_access_token|__proto__|constructor|prototype)"\s*:/i.test(raw)) throw new Error('不能传入凭证、请求头或原型字段；身份由网关提供');
  const args = z.fromJSONSchema(op.input as Parameters<typeof z.fromJSONSchema>[0]).parse(input.arguments) as OpenApiInput['arguments'];
  const url = op.path.replace(/:([a-zA-Z_][\w]*)/g, (_all, key: string) => {
    const value = args.path?.[key];
    if (!value || value === '.' || value === '..' || /[%\\/\x00-\x1f]/.test(value)) throw new Error(`路径参数 ${key} 无效`);
    return encodeURIComponent(value);
  });
  if (!url.startsWith('/open-apis/') || url.includes('://') || url.includes('?') || url.includes('#')) throw new Error('只能调用配置的飞书域名下的业务 API');
  return { method: op.method, url, params: args.params, data: args.data };
}
export async function executeOpenApi(api: OfficeApi, input: OpenApiInput, token?: string) {
  if (input.permissionMode === 'user' && !token) throw new Error('缺少请求者用户授权');
  const result = await api.call({ ...prepareApiRequest(input), userAccessToken: input.permissionMode === 'user' ? token : undefined });
  const text = JSON.stringify(result);
  return text.length <= 40_000 ? result : { truncated: true, preview: text.slice(0, 30_000), has_more: result.has_more, page_token: result.page_token, note: '结果过大，请缩小范围或分页。' };
}
export function openApiSummary(input: OpenApiInput) {
  const request = prepareApiRequest(input);
  return `${input.summary}\n身份：${input.permissionMode === 'user' ? '请求者用户' : '应用'}\n${request.method} ${request.url}\n参数：${JSON.stringify({ params: request.params, data: request.data })}`;
}
