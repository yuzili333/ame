import { describe, expect, it } from 'vitest';
import { OfficeApi, type ApiRequest } from '../office/client';
import { ReleaseMaterialsService } from '../office/release-materials';

const file = { token: 'shtSecretToken', name: '版本清单', type: 'sheet' as const };

function colIndex(label: string): number {
  return [...label].reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function sheetService(sheets: Array<Record<string, unknown>>, cells: Record<string, Record<string, unknown>>) {
  const calls: ApiRequest[] = [];
  const api = new OfficeApi({
    request: async request => {
      calls.push(request);
      if (request.url.endsWith('/sheets/query')) return { code: 0, data: { sheets } };
      const range = decodeURIComponent(request.url.split('/values/')[1] ?? '');
      const match = /^([^!]+)!([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
      if (!match) throw new Error(`Unexpected range ${range}`);
      const [, sheetId, startCol, startRow, endCol, endRow] = match;
      const matrix: unknown[][] = [];
      for (let row = Number(startRow); row <= Number(endRow); row++) {
        const values: unknown[] = [];
        for (let col = colIndex(startCol); col <= colIndex(endCol); col++) {
          let n = col + 1;
          let label = '';
          while (n > 0) { label = String.fromCharCode(65 + (n - 1) % 26) + label; n = Math.floor((n - 1) / 26); }
          values.push(cells[sheetId]?.[`${label}${row}`] ?? '');
        }
        matrix.push(values);
      }
      return { code: 0, data: { valueRange: { values: matrix } } };
    },
    send: async () => undefined,
  });
  return { service: new ReleaseMaterialsService(api), calls };
}

describe('release Sheets preflight', () => {
  it('accepts different field subsets per worksheet and bounds each read', async () => {
    const sheets = [
      { sheet_id: 'empty', title: '说明', grid_properties: { row_count: 2, column_count: 2 }, merges: [] },
      { sheet_id: 'wide', title: '明细', grid_properties: { row_count: 10, column_count: 250 }, merges: [] },
      { sheet_id: 'owner', title: '责任人', grid_properties: { row_count: 1, column_count: 1 }, merges: [] },
    ];
    const { service, calls } = sheetService(sheets, { wide: { IP10: '{{下发编号}}' }, owner: { A1: '{{版本负责人}}' } });
    const result = await service.scanFile(file, { 下发编号: '20261030', 版本负责人: 'Ana', 下发日期: '20261031' });
    expect(result.patches.map(p => [p.location, p.after])).toEqual([['明细!IP10', '20261030'], ['责任人!A1', 'Ana']]);
    const reads = calls.filter(call => call.url.includes('/values/'));
    expect(reads.length).toBeGreaterThan(2);
    for (const call of reads) {
      expect(call.url).toContain('!');
      expect(call.url).toContain(':');
      const match = /!([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(decodeURIComponent(call.url));
      expect(match).toBeTruthy();
      expect((colIndex(match![3]) - colIndex(match![1]) + 1) * (Number(match![4]) - Number(match![2]) + 1)).toBeLessThanOrEqual(1000);
    }
  });

  it('records a merged placeholder at its anchor and rejects a non-anchor marker', async () => {
    const sheets = [{ sheet_id: 'sheet1', title: '封面', grid_properties: { row_count: 2, column_count: 3 }, merges: [{ start_row_index: 0, end_row_index: 0, start_column_index: 0, end_column_index: 2 }] }];
    const { service } = sheetService(sheets, { sheet1: { A1: '{{下发编号}}' } });
    const result = await service.scanFile(file, { 下发编号: '20261030' });
    expect(result.patches[0]).toMatchObject({ location: '封面!A1:C1', cell: 'A1', mergeRange: 'A1:C1' });
    const invalid = sheetService(sheets, { sheet1: { B1: '{{下发编号}}' } });
    await expect(invalid.service.scanFile(file, { 下发编号: '20261030' })).rejects.toThrow('非左上角');
  });

  it('accepts an inclusive one-cell merge and leaves unmatched placeholders unchanged', async () => {
    const sheets = [{ sheet_id: 'sheet1', title: '封面', grid_properties: { row_count: 1, column_count: 3 }, merges: [
      { start_row_index: 0, end_row_index: 0, start_column_index: 0, end_column_index: 0 },
      { start_row_index: 0, end_row_index: 0, start_column_index: 1, end_column_index: 2 },
    ] }];
    const { service } = sheetService(sheets, { sheet1: { A1: '{{下发编号}}', C1: '{{其他字段}}' } });
    const result = await service.scanFile(file, { 下发编号: '20261030' });
    expect(result.patches).toHaveLength(1);
    expect(result.patches[0]).toMatchObject({ cell: 'A1', mergeRange: 'A1:A1', after: '20261030' });
  });

  it('replaces matched markers inside ordinary mixed text, including a merged anchor', async () => {
    const sheets = [{ sheet_id: 'sheet1', title: '封面', grid_properties: { row_count: 1, column_count: 3 }, merges: [
      { start_row_index: 0, end_row_index: 0, start_column_index: 0, end_column_index: 1 },
    ] }];
    const { service } = sheetService(sheets, { sheet1: {
      A1: 'L_REL_{{下发编号}}',
      B1: '{{未匹配字段}}',
      C1: '负责人：{{版本负责人}}；日期：{{下发日期}}；其他：{{未匹配字段}}',
    } });
    const result = await service.scanFile(file, { 下发编号: '202610301709', 版本负责人: 'Ana', 下发日期: '20261031' });
    expect(result.patches.map(patch => [patch.cell, patch.after])).toEqual([
      ['A1', 'L_REL_202610301709'],
      ['C1', '负责人：Ana；日期：20261031；其他：{{未匹配字段}}'],
    ]);
    expect(result.patches[0].mergeRange).toBe('A1:B1');
  });

  it('still rejects formulas containing a matched marker when rendered as a string', async () => {
    const sheets = [{ sheet_id: 'sheet1', title: '封面', grid_properties: { row_count: 1, column_count: 1 }, merges: [] }];
    const { service } = sheetService(sheets, { sheet1: { A1: '=CONCAT("L_REL_{{下发编号}}")' } });
    await expect(service.scanFile(file, { 下发编号: '202610301709' })).rejects.toThrow('公式');
  });

  it('uses an explicit worksheet range instead of scanning a large allocated grid', async () => {
    const sheets = [{ sheet_id: 'large', title: '计划', grid_properties: { row_count: 20000, column_count: 300 }, merges: [] }];
    const { service, calls } = sheetService(sheets, { large: { B2: '{{下发编号}}', C3: '{{版本负责人}}' } });
    const result = await service.scanFile(file, { 下发编号: '20261030', 版本负责人: 'Ana' }, new Set(), [{ fileName: '版本清单', sheetName: '计划', range: 'A1:B2' }]);
    expect(result.patches.map(p => p.cell)).toEqual(['B2']);
    expect(calls.filter(call => call.url.includes('/values/'))).toHaveLength(1);
    await expect(service.scanFile(file, { 下发编号: '20261030' })).rejects.toThrow('扫描范围过大');
  });

  it('preserves the failing endpoint and field violation without exposing resource tokens', async () => {
    const api = new OfficeApi({
      request: async () => ({ code: 99992402, msg: 'field validation failed', error: { log_id: 'log-123', field_violations: [{ field: 'range', description: 'invalid range' }] } }),
      send: async () => undefined,
    });
    const request = api.call({ method: 'GET', url: '/open-apis/sheets/v2/spreadsheets/shtSecretToken/values/sheet1%21A1%3AC1' });
    await expect(request).rejects.toThrow(/99992402.*GET .*:spreadsheet_token.*range.*invalid range.*log-123/);
    await expect(request).rejects.not.toThrow('shtSecretToken');
  });

  it('adds endpoint context when the SDK throws only a generic field validation error', async () => {
    const api = new OfficeApi({ request: async () => { throw new Error('Feishu API 99992402: field validation failed'); }, send: async () => undefined });
    await expect(api.call({ method: 'GET', url: '/open-apis/sheets/v3/spreadsheets/shtSecretToken/sheets/query' }))
      .rejects.toThrow(/99992402.*GET \/open-apis\/sheets\/v3\/spreadsheets\/:spreadsheet_token\/sheets\/query/);
  });
});
