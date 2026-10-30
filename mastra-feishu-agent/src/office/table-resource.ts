import { z } from 'zod';
import { config } from '../config';
import type { OfficeApi } from './client';

const key = z.string().regex(/^[\w@.-]{1,256}$/);
export type TableResource =
  | { type: 'sheet'; token: string; sheetId?: string }
  | { type: 'bitable'; token: string; tableId?: string }
  | { type: 'docx'; token: string };

function allowedHost(hostname: string): boolean {
  const hosts = config.documentHosts;
  return hosts.some(host => hostname === host || hostname.endsWith(`.${host}`));
}

export async function resolveTableResource(input: string, api: OfficeApi, userAccessToken?: string): Promise<TableResource> {
  let url: URL;
  try { url = new URL(input); } catch { throw new Error('请提供完整的飞书表格或 Wiki 链接'); }
  if (!['https:', 'http:'].includes(url.protocol) || !allowedHost(url.hostname)) throw new Error('链接域名不在 FEISHU_DOCUMENT_HOSTS 中');
  const parts = url.pathname.split('/').filter(Boolean);
  const kind = parts[0];
  const raw = parts[1];
  if (!raw) throw new Error('链接中缺少云文档 token');
  const token = key.parse(raw);
  if (kind === 'wiki') {
    const data = await api.call({ method: 'GET', url: '/open-apis/wiki/v2/spaces/get_node', params: { token, obj_type: 'wiki' }, userAccessToken });
    const node = data.node as { obj_type?: string; obj_token?: string } | undefined;
    if (!node?.obj_token) throw new Error('Wiki 节点没有返回真实资源 token');
    const objToken = key.parse(node.obj_token);
    if (node.obj_type === 'sheet') return { type: 'sheet', token: objToken, sheetId: url.searchParams.get('sheet') ?? undefined };
    if (node.obj_type === 'bitable') return { type: 'bitable', token: objToken, tableId: url.searchParams.get('table') ?? undefined };
    if (node.obj_type === 'docx') return { type: 'docx', token: objToken };
    throw new Error(`暂不支持此 Wiki 节点类型：${node.obj_type ?? '未知'}`);
  }
  if (kind === 'sheets') return { type: 'sheet', token, sheetId: url.searchParams.get('sheet') ?? undefined };
  if (kind === 'base') return { type: 'bitable', token, tableId: url.searchParams.get('table') ?? undefined };
  if (kind === 'docx') return { type: 'docx', token };
  throw new Error(`不支持的飞书文档链接类型：${kind}`);
}
