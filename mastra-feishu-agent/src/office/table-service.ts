import { z } from 'zod';
import type { OfficeAction } from './actions';
import type { OfficeApi } from './client';
import { resolveTableResource, type TableResource } from './table-resource';

const scalar = z.union([z.string().max(40_000), z.number(), z.boolean()]);
const fieldMap = z.record(z.string().min(1), scalar).refine(v => Object.keys(v).length > 0);
const common = { url: z.string().url(), permissionMode: z.enum(['app', 'user']).optional() };
export const tableIntent = z.discriminatedUnion('operation', [
  z.object({ ...common, operation: z.literal('inspect'), sheetId: z.string().optional(), tableId: z.string().optional(), range: z.string().optional() }),
  z.object({ ...common, operation: z.literal('fill_first_empty'), sheetId: z.string().optional(), columnHeader: z.string().min(1), value: scalar }),
  z.object({ ...common, operation: z.literal('append_row'), sheetId: z.string().optional(), fields: fieldMap }),
  z.object({ ...common, operation: z.literal('update_cell'), sheetId: z.string().optional(), cell: z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/).optional(), matchColumnHeader: z.string().optional(), matchValue: scalar.optional(), targetColumnHeader: z.string().optional(), value: scalar }),
  z.object({ ...common, operation: z.literal('create_record'), tableId: z.string().optional(), fields: fieldMap }),
  z.object({ ...common, operation: z.literal('update_record'), tableId: z.string().optional(), recordId: z.string().optional(), matchField: z.string().optional(), matchValue: scalar.optional(), fields: fieldMap }),
]);
export type TableIntent = z.infer<typeof tableIntent>;
type Sheet = { sheet_id?: string; title?: string; grid_properties?: { row_count?: number; column_count?: number } };
type Table = { table_id?: string; name?: string };
type BitableRecord = { record_id?: string; fields?: Record<string, unknown> };

const enc = encodeURIComponent;
function column(index: number): string {
  let result = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
  return result;
}
function same(a: unknown, b: unknown): boolean { return JSON.stringify(a ?? null) === JSON.stringify(b ?? null); }
function display(v: unknown): string { return typeof v === 'string' ? v : String(v ?? ''); }
function matchesValue(actual: unknown, expected: string | number | boolean): boolean {
  if (same(actual, expected)) return true;
  if (typeof expected === 'string' && Array.isArray(actual)) return actual.map(v => typeof v === 'object' && v !== null && 'text' in v ? String(v.text) : '').join('') === expected;
  return false;
}

export class TableService {
  constructor(private readonly api: OfficeApi) {}

  async resolve(url: string, userAccessToken?: string) { return resolveTableResource(url, this.api, userAccessToken); }

  private async assertWritable(resource: Extract<TableResource, { type: 'sheet' | 'bitable' }>, userAccessToken?: string) {
    const data = await this.api.call({ method: 'GET', url: `/open-apis/drive/v1/permissions/${enc(resource.token)}/members/auth`, params: { type: resource.type === 'sheet' ? 'sheet' : 'bitable', action: 'edit' }, userAccessToken });
    if (data.auth_result !== true) throw new Error('目标资源没有编辑权限；请由所有者将当前应用或用户设为可编辑协作者');
  }

  private async sheets(token: string, userAccessToken?: string): Promise<Sheet[]> {
    const data = await this.api.call({ method: 'GET', url: `/open-apis/sheets/v3/spreadsheets/${enc(token)}/sheets/query`, userAccessToken });
    return (data.sheets as Sheet[] | undefined) ?? [];
  }
  private async sheet(resource: Extract<TableResource, { type: 'sheet' }>, preferred?: string, userAccessToken?: string): Promise<Sheet & { sheet_id: string }> {
    const sheets = await this.sheets(resource.token, userAccessToken);
    const choice = preferred ?? resource.sheetId;
    const matches = choice ? sheets.filter(s => s.sheet_id === choice || s.title === choice) : sheets;
    if (matches.length !== 1 || !matches[0].sheet_id) throw new Error(`无法唯一确定工作表；可选：${sheets.map(s => `${s.title ?? '?'} (${s.sheet_id ?? '?'})`).join('、')}`);
    return matches[0] as Sheet & { sheet_id: string };
  }
  private async values(token: string, sheetId: string, range: string, userAccessToken?: string): Promise<unknown[][]> {
    const data = await this.api.call({ method: 'GET', url: `/open-apis/sheets/v2/spreadsheets/${enc(token)}/values/${enc(`${sheetId}!${range}`)}`, userAccessToken });
    const valueRange = data.valueRange as { values?: unknown[][] } | undefined;
    return valueRange?.values ?? [];
  }
  private async header(token: string, sheetId: string, userAccessToken?: string): Promise<string[]> {
    const row = (await this.values(token, sheetId, 'A1:IV1', userAccessToken))[0] ?? [];
    const headers = row.map(display);
    while (headers.at(-1) === '') headers.pop();
    if (!headers.length) throw new Error('工作表首行没有表头');
    return headers;
  }
  private headerIndex(headers: string[], name: string): number {
    const matches = headers.flatMap((v, i) => v === name ? [i] : []);
    if (matches.length !== 1) throw new Error(`表头“${name}”${matches.length ? '重复' : '不存在'}，请提供明确列或修正表头`);
    return matches[0];
  }
  private async columnCells(token: string, sheetId: string, col: number, userAccessToken?: string, stopAtFirstEmpty = false) {
    const letter = column(col);
    const cells: unknown[] = [];
    for (let start = 2; start <= 10_000; start += 500) {
      const end = Math.min(start + 499, 10_000);
      const rows = await this.values(token, sheetId, `${letter}${start}:${letter}${end}`, userAccessToken);
      for (let i = 0; i < end - start + 1; i++) {
        const value = rows[i]?.[0] ?? null;
        cells.push(value);
        if (stopAtFirstEmpty && (value === null || value === '')) return cells;
      }
    }
    return cells;
  }
  private async tables(token: string, userAccessToken?: string): Promise<Table[]> {
    const items: Table[] = [];
    let pageToken: string | undefined;
    do {
      const data = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(token)}/tables`, params: { page_size: 100, ...(pageToken ? { page_token: pageToken } : {}) }, userAccessToken });
      items.push(...((data.items as Table[] | undefined) ?? []));
      pageToken = data.has_more ? data.page_token as string | undefined : undefined;
      if (items.length >= 1000) throw new Error('数据表过多，请提供 tableId');
    } while (pageToken);
    return items;
  }
  private async table(resource: Extract<TableResource, { type: 'bitable' }>, preferred?: string, userAccessToken?: string): Promise<Table & { table_id: string }> {
    const tables = await this.tables(resource.token, userAccessToken);
    const choice = preferred ?? resource.tableId;
    const matches = choice ? tables.filter(t => t.table_id === choice || t.name === choice) : tables;
    if (matches.length !== 1 || !matches[0].table_id) throw new Error(`无法唯一确定数据表；可选：${tables.map(t => `${t.name ?? '?'} (${t.table_id ?? '?'})`).join('、')}`);
    return matches[0] as Table & { table_id: string };
  }
  private async fields(token: string, tableId: string, userAccessToken?: string): Promise<Array<{ field_name?: string; type?: number }>> {
    const items: Array<{ field_name?: string; type?: number }> = [];
    let pageToken: string | undefined;
    do {
      const data = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(token)}/tables/${enc(tableId)}/fields`, params: { page_size: 100, ...(pageToken ? { page_token: pageToken } : {}) }, userAccessToken });
      items.push(...((data.items as typeof items | undefined) ?? []));
      pageToken = data.has_more ? data.page_token as string | undefined : undefined;
      if (items.length >= 1000) throw new Error('字段过多，请缩小操作范围');
    } while (pageToken);
    return items;
  }
  private async records(token: string, tableId: string, userAccessToken?: string): Promise<BitableRecord[]> {
    const items: BitableRecord[] = [];
    let pageToken: string | undefined;
    do {
      const data = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(token)}/tables/${enc(tableId)}/records`, params: { page_size: 100, ...(pageToken ? { page_token: pageToken } : {}) }, userAccessToken });
      items.push(...((data.items as BitableRecord[] | undefined) ?? []));
      pageToken = data.has_more ? data.page_token as string | undefined : undefined;
      if (items.length >= 10_000) throw new Error('记录超过 10000 条，请提供 recordId');
    } while (pageToken);
    return items;
  }
  private async record(token: string, tableId: string, recordId: string, userAccessToken?: string): Promise<BitableRecord> {
    const data = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(token)}/tables/${enc(tableId)}/records/${enc(recordId)}`, userAccessToken });
    return data.record as BitableRecord;
  }
  private validateFields(available: Array<{ field_name?: string; type?: number }>, values: Record<string, string | number | boolean>) {
    for (const [name, value] of Object.entries(values)) {
      const matches = available.filter(f => f.field_name === name);
      if (matches.length !== 1) throw new Error(`多维表格字段“${name}”不存在或不唯一`);
      const type = matches[0].type;
      if (!((type === 1 && typeof value === 'string') || (type === 2 && typeof value === 'number') || (type === 7 && typeof value === 'boolean'))) {
        throw new Error(`字段“${name}”暂不支持此值类型；当前支持文本、数字和复选框`);
      }
    }
  }

  async inspect(intent: Extract<TableIntent, { operation: 'inspect' }>, userAccessToken?: string) {
    const resource = await this.resolve(intent.url, userAccessToken);
    if (resource.type === 'docx') throw new Error('该链接是新版文档，不是电子表格或多维表格');
    if (resource.type === 'sheet') {
      const sheets = await this.sheets(resource.token, userAccessToken);
      if (!intent.sheetId && !resource.sheetId && sheets.length !== 1) return { type: 'sheet', sheets };
      const sheet = await this.sheet(resource, intent.sheetId, userAccessToken);
      const range = intent.range ?? 'A1:J10';
      const match = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/.exec(range);
      const colNumber = (letters: string) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
      if (!match || Number(match[4]) < Number(match[2]) || Number(match[4]) - Number(match[2]) >= 100 || colNumber(match[3]) < colNumber(match[1]) || colNumber(match[3]) - colNumber(match[1]) >= 50) throw new Error('读取范围最多 100 行、50 列，例如 A1:J10');
      return { type: 'sheet', sheets, sheet, range, values: await this.values(resource.token, sheet.sheet_id, range, userAccessToken) };
    }
    const tables = await this.tables(resource.token, userAccessToken);
    if (!intent.tableId && !resource.tableId && tables.length !== 1) return { type: 'bitable', tables };
    const table = await this.table(resource, intent.tableId, userAccessToken);
    const data = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(resource.token)}/tables/${enc(table.table_id)}/records`, params: { page_size: 10 }, userAccessToken });
    return { type: 'bitable', tables, table, fields: await this.fields(resource.token, table.table_id, userAccessToken), records: data.items ?? [] };
  }

  async prepare(intent: Exclude<TableIntent, { operation: 'inspect' }>, permissionMode: 'app' | 'user', userAccessToken?: string): Promise<OfficeAction> {
    const resource = await this.resolve(intent.url, userAccessToken);
    if (intent.operation === 'fill_first_empty' || intent.operation === 'append_row' || intent.operation === 'update_cell') {
      if (resource.type !== 'sheet') throw new Error('目标不是普通电子表格；多维表格请使用记录工具');
      await this.assertWritable(resource, userAccessToken);
      const sheet = await this.sheet(resource, intent.sheetId, userAccessToken);
      if (intent.operation === 'append_row') {
        const headers = await this.header(resource.token, sheet.sheet_id, userAccessToken);
        if (headers.some(h => !h) || new Set(headers).size !== headers.length) throw new Error('追加行需要非空且唯一的表头');
        if (headers.length > 100 || (sheet.grid_properties?.row_count ?? 1000) > 5_000) throw new Error('飞书追加接口单次最多覆盖 5000 行、100 列；请明确指定写入位置');
        const values = headers.map(header => Object.hasOwn(intent.fields, header) ? intent.fields[header] : null);
        for (const name of Object.keys(intent.fields)) this.headerIndex(headers, name);
        return { kind: 'sheet_row_append', spreadsheetToken: resource.token, sheetId: sheet.sheet_id, headers, values, permissionMode };
      }
      let cell: string;
      if (intent.operation === 'fill_first_empty') {
        const headers = await this.header(resource.token, sheet.sheet_id, userAccessToken);
        const col = this.headerIndex(headers, intent.columnHeader);
        const row = (await this.columnCells(resource.token, sheet.sheet_id, col, userAccessToken, true)).findIndex(v => v === null || v === '');
        if (row < 0) throw new Error('前 10000 行没有空白单元格，请指定明确位置');
        cell = `${column(col)}${row + 2}`;
      } else if (intent.cell) {
        cell = intent.cell;
      } else {
        if (!intent.matchColumnHeader || intent.matchValue === undefined || !intent.targetColumnHeader) throw new Error('请提供目标单元格，或匹配列、匹配值和目标列');
        const headers = await this.header(resource.token, sheet.sheet_id, userAccessToken);
        const matchCol = this.headerIndex(headers, intent.matchColumnHeader);
        const targetCol = this.headerIndex(headers, intent.targetColumnHeader);
        const matches = (await this.columnCells(resource.token, sheet.sheet_id, matchCol, userAccessToken)).flatMap((v, i) => same(v, intent.matchValue) ? [i + 2] : []);
        if (matches.length !== 1) throw new Error(`匹配行数量为 ${matches.length}，请提供明确单元格`);
        cell = `${column(targetCol)}${matches[0]}`;
      }
      const before = (await this.values(resource.token, sheet.sheet_id, `${cell}:${cell}`, userAccessToken))[0]?.[0] ?? null;
      if (intent.operation === 'fill_first_empty' && before !== null && before !== '') throw new Error('目标单元格已被填写，请重新发起');
      if (before !== null && before !== '' && typeof before !== 'string' && typeof before !== 'number' && typeof before !== 'boolean') throw new Error('目标单元格含复杂值，暂不支持覆盖');
      return { kind: 'sheet_cell_set', spreadsheetToken: resource.token, sheetId: sheet.sheet_id, cell, expectedValue: before === '' ? null : before as string | number | boolean | null, value: intent.value, permissionMode };
    }
    if (resource.type !== 'bitable') throw new Error('目标不是多维表格；普通电子表格请使用单元格工具');
    await this.assertWritable(resource, userAccessToken);
    const table = await this.table(resource, intent.tableId, userAccessToken);
    const availableFields = await this.fields(resource.token, table.table_id, userAccessToken);
    this.validateFields(availableFields, intent.fields);
    if (intent.operation === 'create_record') return { kind: 'bitable_record_create', appToken: resource.token, tableId: table.table_id, fields: intent.fields, permissionMode };
    let record: BitableRecord;
    if (intent.recordId) record = await this.record(resource.token, table.table_id, intent.recordId, userAccessToken);
    else {
      if (!intent.matchField || intent.matchValue === undefined) throw new Error('请提供 recordId，或用于唯一匹配的字段和值');
      if (availableFields.filter(field => field.field_name === intent.matchField).length !== 1) throw new Error(`匹配字段“${intent.matchField}”不存在或不唯一`);
      const matches = (await this.records(resource.token, table.table_id, userAccessToken)).filter(r => matchesValue(r.fields?.[intent.matchField!], intent.matchValue!));
      if (matches.length !== 1) throw new Error(`匹配记录数量为 ${matches.length}，请提供 recordId`);
      record = matches[0];
    }
    if (!record.record_id) throw new Error('记录缺少 recordId');
    const expectedFields = Object.fromEntries(Object.keys(intent.fields).map(name => [name, record.fields?.[name] ?? null]));
    return { kind: 'bitable_record_update', appToken: resource.token, tableId: table.table_id, recordId: record.record_id, fields: intent.fields, expectedFields, permissionMode };
  }

  async execute(action: Extract<OfficeAction, { kind: 'sheet_cell_set' | 'sheet_row_append' | 'bitable_record_create' | 'bitable_record_update' }>, userAccessToken?: string) {
    if (action.permissionMode === 'user' && !userAccessToken) throw new Error('用户表格授权已失效，请重新授权');
    if (action.kind === 'sheet_cell_set') {
      const before = (await this.values(action.spreadsheetToken, action.sheetId, `${action.cell}:${action.cell}`, userAccessToken))[0]?.[0] ?? null;
      if (!same(before === '' ? null : before, action.expectedValue)) throw new Error('审批后单元格内容已变化，未执行写入；请重新发起');
      return this.api.call({ method: 'PUT', url: `/open-apis/sheets/v2/spreadsheets/${enc(action.spreadsheetToken)}/values`, data: { valueRange: { range: `${action.sheetId}!${action.cell}:${action.cell}`, values: [[action.value]] } }, userAccessToken });
    }
    if (action.kind === 'sheet_row_append') {
      if (action.headers.length !== action.values.length) throw new Error('表头与追加行长度不一致');
      const sheet = await this.sheet({ type: 'sheet', token: action.spreadsheetToken, sheetId: action.sheetId }, undefined, userAccessToken);
      const rowCount = sheet.grid_properties?.row_count ?? 1000;
      if (action.headers.length > 100 || rowCount > 5_000) throw new Error('飞书追加接口单次最多覆盖 5000 行、100 列，未执行追加');
      const headers = await this.header(action.spreadsheetToken, action.sheetId, userAccessToken);
      if (!same(headers, action.headers)) throw new Error('审批后表头已变化，未执行追加；请重新发起');
      return this.api.call({ method: 'POST', url: `/open-apis/sheets/v2/spreadsheets/${enc(action.spreadsheetToken)}/values_append`, params: { insertDataOption: 'INSERT_ROWS' }, data: { valueRange: { range: `${action.sheetId}!A1:${column(action.headers.length - 1)}${rowCount}`, values: [action.values] } }, userAccessToken });
    }
    const base = `/open-apis/bitable/v1/apps/${enc(action.appToken)}/tables/${enc(action.tableId)}/records`;
    if (action.kind === 'bitable_record_create') return this.api.call({ method: 'POST', url: base, data: { fields: action.fields }, userAccessToken });
    const record = await this.record(action.appToken, action.tableId, action.recordId, userAccessToken);
    for (const [name, expected] of Object.entries(action.expectedFields)) {
      if (!same(record.fields?.[name], expected)) throw new Error(`审批后字段“${name}”已变化，未执行写入；请重新发起`);
    }
    return this.api.call({ method: 'PUT', url: `${base}/${enc(action.recordId)}`, data: { fields: action.fields }, userAccessToken });
  }
}
