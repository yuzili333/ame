import { describe, expect, it } from 'vitest';
import { OfficeApi, type ApiRequest } from '../office/client';
import { resolveTableResource } from '../office/table-resource';
import { TableService } from '../office/table-service';

function mockApi(handler: (request: ApiRequest) => Record<string, unknown>) {
  const requests: ApiRequest[] = [];
  const api = new OfficeApi({
    request: async request => { requests.push(request); return { code: 0, data: handler(request) }; },
    send: async () => undefined,
  });
  return { api, requests };
}

function sheetApi() {
  return mockApi(request => {
    if (request.url.includes('/members/auth')) return { auth_result: true };
    if (request.url.includes('/sheets/query')) return { sheets: [{ sheet_id: 's1', title: 'Main', grid_properties: { row_count: 1000 } }] };
    if (request.url.includes('/values/')) {
      const range = decodeURIComponent(request.url.split('/values/')[1]);
      if (range === 's1!A1:IV1') return { valueRange: { values: [['名称', '状态']] } };
      if (range === 's1!A2:A501') return { valueRange: { values: [['Existing'], []] } };
      if (range === 's1!A3:A3') return { valueRange: { values: [] } };
      if (range === 's1!B2:B2') return { valueRange: { values: [['open']] } };
      return { valueRange: { values: [] } };
    }
    return { ok: true };
  });
}

describe('Feishu table resource routing', () => {
  it('resolves a Wiki node to its real spreadsheet token without calling Docx', async () => {
    const { api, requests } = mockApi(() => ({ node: { obj_type: 'sheet', obj_token: 'sht-real' } }));
    expect(await resolveTableResource('https://tenant.feishu.cn/wiki/wiki-node?sheet=s1', api)).toEqual({ type: 'sheet', token: 'sht-real', sheetId: 's1' });
    expect(requests).toMatchObject([{ method: 'GET', url: '/open-apis/wiki/v2/spaces/get_node', params: { token: 'wiki-node', obj_type: 'wiki' } }]);
  });

  it('routes direct Base and Sheets URLs and rejects an unrelated host', async () => {
    const { api, requests } = mockApi(() => ({}));
    expect(await resolveTableResource('https://tenant.feishu.cn/base/app1?table=tbl1', api)).toEqual({ type: 'bitable', token: 'app1', tableId: 'tbl1' });
    expect(await resolveTableResource('https://tenant.feishu.cn/sheets/sht1', api)).toEqual({ type: 'sheet', token: 'sht1' });
    await expect(resolveTableResource('https://example.invalid/wiki/x', api)).rejects.toThrow('FEISHU_DOCUMENT_HOSTS');
    expect(requests).toHaveLength(0);
  });

  it('reports unsupported Wiki types instead of using Docx', async () => {
    const { api, requests } = mockApi(() => ({ node: { obj_type: 'file', obj_token: 'file1' } }));
    await expect(resolveTableResource('https://tenant.feishu.cn/wiki/node1', api)).rejects.toThrow('暂不支持');
    expect(requests.some(r => r.url.includes('/docx/'))).toBe(false);
  });

  it('never sends a Wiki token to the Docx update endpoint', async () => {
    const { api, requests } = mockApi(() => ({}));
    await expect(api.execute({ kind: 'doc_update', documentId: 'wikcn123', blockId: 'block1', text: 'x' })).rejects.toThrow('不能作为 Docx');
    expect(requests).toHaveLength(0);
  });

  it('rejects a bare Wiki node token in the Mastra document reader before calling Feishu', async () => {
    const { readDocumentTool } = await import('../mastra/tools/office-tools');
    const { setOfficeApi } = await import('../runtime/office-bridge');
    const { api, requests } = mockApi(() => ({}));
    setOfficeApi(api);
    try {
      await expect(readDocumentTool.execute!({ documentId: 'A1dXwT0x1ilfxikmDHhcT0wPnve' }, {} as never)).rejects.toThrow('Wiki 节点 token');
      expect(requests).toHaveLength(0);
    } finally { setOfficeApi(undefined); }
  });
});

describe('Feishu spreadsheet edits', () => {
  it('finds the first empty cell under a unique header and checks it again before writing', async () => {
    const { api, requests } = sheetApi();
    const service = new TableService(api);
    const action = await service.prepare({ operation: 'fill_first_empty', url: 'https://tenant.feishu.cn/sheets/sht1', columnHeader: '名称', value: 'Topson' }, 'app');
    expect(action).toMatchObject({ kind: 'sheet_cell_set', spreadsheetToken: 'sht1', sheetId: 's1', cell: 'A3', expectedValue: null, value: 'Topson' });
    await api.execute(action);
    expect(requests.find(r => r.method === 'PUT')).toMatchObject({ url: '/open-apis/sheets/v2/spreadsheets/sht1/values', data: { valueRange: { range: 's1!A3:A3', values: [['Topson']] } } });
  });

  it('stops if a cell changed after approval', async () => {
    const { api, requests } = mockApi(request => {
      if (request.url.includes('/values/')) return { valueRange: { values: [['someone else']] } };
      return {};
    });
    await expect(api.execute({ kind: 'sheet_cell_set', spreadsheetToken: 'sht1', sheetId: 's1', cell: 'A2', expectedValue: null, value: 'Topson', permissionMode: 'app' })).rejects.toThrow('已变化');
    expect(requests.some(r => r.method === 'PUT')).toBe(false);
  });

  it('appends a row with INSERT_ROWS and does not overwrite existing data', async () => {
    const { api, requests } = sheetApi();
    const service = new TableService(api);
    const action = await service.prepare({ operation: 'append_row', url: 'https://tenant.feishu.cn/sheets/sht1', fields: { 名称: 'Topson', 状态: 'open' } }, 'app');
    expect(action).toMatchObject({ kind: 'sheet_row_append', headers: ['名称', '状态'], values: ['Topson', 'open'] });
    await api.execute(action);
    expect(requests.find(r => r.url.endsWith('/values_append'))).toMatchObject({ method: 'POST', params: { insertDataOption: 'INSERT_ROWS' }, data: { valueRange: { range: 's1!A1:B1000', values: [['Topson', 'open']] } } });
  });

  it('rejects an append range larger than the Feishu API limit before approval', async () => {
    const { api, requests } = mockApi(request => {
      if (request.url.includes('/members/auth')) return { auth_result: true };
      if (request.url.includes('/sheets/query')) return { sheets: [{ sheet_id: 's1', title: 'Large', grid_properties: { row_count: 5001 } }] };
      return { valueRange: { values: [['名称']] } };
    });
    await expect(new TableService(api).prepare({ operation: 'append_row', url: 'https://tenant.feishu.cn/sheets/sht1?sheet=s1', fields: { 名称: 'Topson' } }, 'app')).rejects.toThrow('5000 行');
    expect(requests.some(r => r.url.endsWith('/values_append'))).toBe(false);
  });

  it('asks for a worksheet when a link does not select one', async () => {
    const { api } = mockApi(request => request.url.includes('/members/auth') ? { auth_result: true } : { sheets: [{ sheet_id: 'a', title: 'A' }, { sheet_id: 'b', title: 'B' }] });
    await expect(new TableService(api).prepare({ operation: 'fill_first_empty', url: 'https://tenant.feishu.cn/sheets/sht1', columnHeader: '名称', value: 'x' }, 'app')).rejects.toThrow('无法唯一确定工作表');
  });

  it('updates a cell by uniquely matched row', async () => {
    const { api } = sheetApi();
    const action = await new TableService(api).prepare({ operation: 'update_cell', url: 'https://tenant.feishu.cn/sheets/sht1', matchColumnHeader: '名称', matchValue: 'Existing', targetColumnHeader: '状态', value: 'done' }, 'app');
    expect(action).toMatchObject({ kind: 'sheet_cell_set', cell: 'B2', expectedValue: 'open', value: 'done' });
  });

  it('allows an explicit A1 cell even when the sheet has no header row', async () => {
    const { api } = mockApi(request => {
      if (request.url.includes('/members/auth')) return { auth_result: true };
      if (request.url.includes('/sheets/query')) return { sheets: [{ sheet_id: 's1', title: 'Blank' }] };
      return { valueRange: { values: [] } };
    });
    const action = await new TableService(api).prepare({ operation: 'update_cell', url: 'https://tenant.feishu.cn/sheets/sht1?sheet=s1', cell: 'B7', value: 'ok' }, 'app');
    expect(action).toMatchObject({ kind: 'sheet_cell_set', cell: 'B7', expectedValue: null });
  });
});

describe('Feishu Base record edits', () => {
  function baseApi() {
    return mockApi(request => {
      if (request.url.includes('/members/auth')) return { auth_result: true };
      if (request.url.endsWith('/tables')) return { items: [{ table_id: 'tbl1', name: 'Tasks' }] };
      if (request.url.endsWith('/fields')) return { items: [{ field_name: '名称', type: 1 }, { field_name: '数量', type: 2 }] };
      if (request.url.endsWith('/records')) return request.method === 'GET' ? { items: [{ record_id: 'rec1', fields: { 名称: [{ text: 'Old' }], 数量: 1 } }] } : { record: { record_id: 'rec2' } };
      if (request.url.endsWith('/records/rec1')) return { record: { record_id: 'rec1', fields: { 名称: [{ text: 'Old' }], 数量: 1 } } };
      return {};
    });
  }

  it('creates one record with validated fields', async () => {
    const { api, requests } = baseApi();
    const action = await new TableService(api).prepare({ operation: 'create_record', url: 'https://tenant.feishu.cn/base/app1', fields: { 名称: 'Topson', 数量: 2 } }, 'app');
    expect(action.kind).toBe('bitable_record_create');
    await api.execute(action);
    expect(requests.find(r => r.method === 'POST')).toMatchObject({ url: '/open-apis/bitable/v1/apps/app1/tables/tbl1/records', data: { fields: { 名称: 'Topson', 数量: 2 } } });
  });

  it('updates a uniquely matched record and rechecks prior fields', async () => {
    const { api, requests } = baseApi();
    const action = await new TableService(api).prepare({ operation: 'update_record', url: 'https://tenant.feishu.cn/base/app1', matchField: '名称', matchValue: 'Old', fields: { 数量: 3 } }, 'app');
    expect(action).toMatchObject({ kind: 'bitable_record_update', recordId: 'rec1', expectedFields: { 数量: 1 } });
    await api.execute(action);
    expect(requests.find(r => r.method === 'PUT')).toMatchObject({ url: '/open-apis/bitable/v1/apps/app1/tables/tbl1/records/rec1', data: { fields: { 数量: 3 } } });
  });

  it('rejects unsupported field types before proposing a write', async () => {
    const { api, requests } = baseApi();
    await expect(new TableService(api).prepare({ operation: 'create_record', url: 'https://tenant.feishu.cn/base/app1', fields: { 数量: 'two' } }, 'app')).rejects.toThrow('暂不支持');
    expect(requests.some(r => r.method === 'POST' || r.method === 'PUT')).toBe(false);
  });
});
