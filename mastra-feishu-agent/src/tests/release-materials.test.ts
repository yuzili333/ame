import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const dataDir = mkdtempSync(join(tmpdir(), 'feishu-release-test-'));
let OfficeApi: typeof import('../office/client').OfficeApi;
let ReleaseMaterialsService: typeof import('../office/release-materials').ReleaseMaterialsService;
let parseDriveFolderUrl: typeof import('../office/release-materials').parseDriveFolderUrl;
let stateDb: typeof import('../state/store').stateDb;

beforeAll(async () => {
  vi.stubEnv('AGENT_DATA_DIR', dataDir);
  vi.resetModules();
  ({ OfficeApi } = await import('../office/client'));
  ({ ReleaseMaterialsService, parseDriveFolderUrl } = await import('../office/release-materials'));
  ({ stateDb } = await import('../state/store'));
});
afterAll(() => { vi.unstubAllEnvs(); rmSync(dataDir, { recursive: true, force: true }); });

const input = {
  templateFolderUrl: 'https://feishu.cn/drive/folder/template',
  materialsParentUrl: 'https://feishu.cn/drive/folder/parent',
  releaseId: '20261029',
  fields: { 负责人: '张三', 日期: '2026-10-29' },
};
const caller = { userId: 'u-release', chatId: 'g-release', chatType: 'group' as const };

async function approvedJob(service: InstanceType<typeof ReleaseMaterialsService>, manifest: import('../office/release-materials').ReleaseManifest) {
  const { jobId } = await service.createJob(caller, manifest);
  const { proposeAction } = await import('../state/approvals');
  const action = await proposeAction(caller, { kind: 'release_materials', jobId, releaseId: manifest.input.releaseId, fileCount: manifest.files.length, fieldSummary: 'test' });
  const db = await stateDb();
  await db.execute({ sql: 'UPDATE release_jobs SET approval_id = ? WHERE id = ?', args: [action.id, jobId] });
  await db.execute({ sql: "UPDATE actions SET status = 'executing' WHERE id = ?", args: [action.id] });
  return jobId;
}

function fixture(options: { unsupported?: boolean; sourceChanged?: boolean; duplicateFolder?: boolean; failCopy?: boolean; failWrite?: boolean; copyReadDelay?: boolean; formulaMarker?: boolean; ownerTypo?: boolean; docTitleMarker?: boolean; sheetTitleMarker?: boolean; sheetMixedText?: boolean; mergedSheet?: boolean; copyMergeChangesAfterScan?: boolean; plainTemplates?: boolean; parentGroupPerm?: 'view' | 'edit'; failFolderForbidden?: boolean } = {}) {
  const calls: import('../office/client').ApiRequest[] = [];
  let created = false;
  let sourceReads = 0;
  const originalDocText = options.plainTemplates ? '普通说明' : options.ownerTypo ? '{{下发编号}} / {{版本负责任}}' : '{{下发编号}} / {{负责人}}';
  const originalSheetText = options.plainTemplates ? '普通表格' : options.sheetMixedText ? 'L_REL_{{下发编号}}' : '{{日期}}';
  const originalBaseText = options.plainTemplates ? '普通记录' : '{{负责人}}';
  let docText = originalDocText;
  let copyTitle = options.docTitleMarker ? '说明{{下发编号}}' : '说明.docx';
  let copySheetName = options.sheetTitleMarker ? '清单{{下发编号}}' : '清单';
  let sheetText = originalSheetText;
  let baseText = originalBaseText;
  const copied = new Set<string>();
  let copyReadFailed = false;
  let copySheetQueries = 0;
  const sourceItems = [
    { token: 'doxcnSource', name: '说明.docx', type: 'docx', modified_time: '1' },
    { token: 'shtSource', name: options.sheetTitleMarker ? '清单{{下发编号}}' : '清单', type: 'sheet', modified_time: '1' },
    { token: 'child', name: '子目录', type: 'folder' },
  ];
  const transport = {
    request: async (request: import('../office/client').ApiRequest) => {
      calls.push(request);
      const url = request.url;
      let data: Record<string, unknown> = {};
      if (url.includes('/permissions/') && url.endsWith('/members')) {
        data = { items: [{ member_type: 'openchat', member_id: caller.chatId, perm: options.parentGroupPerm ?? 'edit' }] };
      } else
      if (url === '/open-apis/drive/v1/files' && request.method === 'GET') {
        const folder = request.params?.folder_token;
        if (folder === 'template') data = request.params?.page_token ? { files: [{ token: options.unsupported ? 'pdf1' : 'baseSource', name: '责任表', type: options.unsupported ? 'file' : 'bitable', modified_time: '1' }], has_more: false } : { files: sourceItems, has_more: true, next_page_token: 'next' };
        else if (folder === 'parent') data = { files: options.duplicateFolder ? [{ token: 'existing', name: '20261029', type: 'folder' }] : created ? [{ token: 'newFolder', name: '20261029', type: 'folder' }] : [], has_more: false };
        else if (folder === 'newFolder') data = { files: [
          ...(copied.has('doxcnCopy') ? [{ token: 'doxcnCopy', name: copyTitle, type: 'docx' }] : []),
          ...(copied.has('shtCopy') ? [{ token: 'shtCopy', name: copySheetName, type: 'sheet' }] : []),
          ...(copied.has('baseCopy') ? [{ token: 'baseCopy', name: '责任表', type: 'bitable' }] : []),
        ], has_more: false };
      } else if (url.endsWith('/create_folder')) {
        if (options.failFolderForbidden) return { code: 1061004, msg: 'forbidden' };
        created = true; data = { token: 'newFolder', url: 'https://feishu.cn/drive/folder/newFolder' };
      } else if (url.endsWith('/copy')) {
        if (options.failCopy && url.includes('shtSource')) throw new Error('copy timeout');
        const source = url.split('/').at(-2);
        data = { file: { token: source === 'doxcnSource' ? 'doxcnCopy' : source === 'shtSource' ? 'shtCopy' : 'baseCopy' } };
        copied.add((data.file as { token: string }).token);
      } else if (url === '/open-apis/drive/v1/files/doxcnCopy' && request.method === 'PATCH') {
        copyTitle = (request.data as { new_title: string }).new_title;
      } else if (url === '/open-apis/drive/v1/files/shtCopy' && request.method === 'PATCH') {
        copySheetName = (request.data as { new_title: string }).new_title;
      } else if (url.includes('/docx/v1/documents/') && url.endsWith('/blocks') && request.method === 'GET') {
        const isCopy = url.includes('doxcnCopy');
        if (isCopy && options.copyReadDelay && !copyReadFailed) { copyReadFailed = true; throw new Error('not ready'); }
        if (!isCopy && options.sourceChanged && ++sourceReads > 1) docText = '{{下发编号}} / {{负责人}} / 变更';
        const items: Array<Record<string, unknown>> = [{ block_id: isCopy ? 'blockCopy' : 'blockSource', block_type: 2, text: { elements: [{ text_run: { content: isCopy ? docText.replace('{{下发编号}}', created ? '20261029' : '{{下发编号}}') : docText } }] } }];
        if (options.docTitleMarker) items.unshift({ block_id: isCopy ? 'doxcnCopy' : 'doxcnSource', block_type: 1, page: { elements: [{ text_run: { content: isCopy ? copyTitle : '说明{{下发编号}}' } }] } });
        data = { items, has_more: false };
        // The copy starts from the original marker text until the batch update below.
        if (isCopy && !docWritten) (items.find(item => item.block_type === 2) as { text: { elements: Array<{ text_run: { content: string } }> } }).text.elements[0].text_run.content = originalDocText;
      } else if (url.endsWith('/blocks/batch_update')) {
        if (options.failWrite) throw new Error('write timeout');
        const req = (request.data as { requests: Array<{ update_text_elements: { elements: Array<{ text_run: { content: string } }> } }> }).requests;
        docText = req[0].update_text_elements.elements[0].text_run.content;
        docWritten = true;
      } else if (url.includes('/docx/v1/documents/') && url.includes('/blocks/blockCopy')) {
        data = { block: { block_id: 'blockCopy', text: { elements: [{ text_run: { content: docText } }] } } };
      } else if (url.endsWith('/blocks/doxcnCopy')) {
        data = { block: { block_id: 'doxcnCopy', block_type: 1, page: { elements: [{ text_run: { content: copyTitle } }] } } };
      } else if (url.includes('/sheets/query')) {
        const isCopy = url.includes('shtCopy');
        if (isCopy) copySheetQueries++;
        const merges = options.mergedSheet && !(isCopy && options.copyMergeChangesAfterScan && copySheetQueries > 1)
          ? [{ start_row_index: 0, end_row_index: 0, start_column_index: 0, end_column_index: 1 }] : [];
        data = { sheets: [{ sheet_id: isCopy ? 'sheetCopy' : 'sheetSource', title: 'Main', grid_properties: { row_count: 1, column_count: 2 }, merges }] };
      } else if (url.includes('/sheets/v2/') && url.includes('/values/') && request.method === 'GET') {
        const isCopy = url.includes('shtCopy');
        data = { valueRange: { values: [[options.formulaMarker ? { type: 'formula', text: '={{日期}}' } : isCopy ? sheetText : originalSheetText, { type: 'formula', text: '=1+1' }]] } };
      } else if (url.includes('/sheets/v2/') && request.method === 'PUT') {
        sheetText = (request.data as { valueRange: { values: string[][] } }).valueRange.values[0][0];
      } else if (url.endsWith('/tables')) {
        data = { items: [{ table_id: url.includes('baseCopy') ? 'tblCopy' : 'tblSource', name: 'Owners' }], has_more: false };
      } else if (url.endsWith('/fields')) {
        data = { items: [{ field_name: '负责人', type: 1 }], has_more: false };
      } else if (url.endsWith('/records')) {
        data = { items: [{ record_id: url.includes('baseCopy') ? 'recCopy' : 'recSource', fields: { 负责人: [{ text: url.includes('baseCopy') ? baseText : originalBaseText }] } }], has_more: false };
      } else if (url.includes('/records/recCopy') && request.method === 'PUT') {
        baseText = (request.data as { fields: { 负责人: string } }).fields.负责人;
      } else if (url.includes('/records/recCopy') && request.method === 'GET') {
        data = { record: { fields: { 负责人: [{ text: baseText }] } } };
      }
      return { code: 0, data };
    },
    send: async () => undefined,
  };
  let docWritten = false;
  return { service: new ReleaseMaterialsService(new OfficeApi(transport)), calls };
}

describe('release materials', () => {
  it('rejects a release ID borrowed from yesterday when the current message names only a new batch', async () => {
    const { requestReleaseMaterials } = await import('../mastra/tools/office-tools');
    const oldInput = { ...input, releaseId: '202610291638' };
    await expect(Reflect.apply(requestReleaseMaterials, undefined, [oldInput, caller, Date.now(), '新增下发版本编号为“202610300856”，版本负责人 Ana']))
      .rejects.toThrow('当前消息');
    const rows = await (await stateDb()).execute({ sql: 'SELECT id FROM release_jobs WHERE manifest_json LIKE ?', args: ['%202610291638%'] });
    expect(rows.rows).toHaveLength(0);
  });
  it('accepts folder links copied from the Feishu client on my.feishu.cn', () => {
    expect(parseDriveFolderUrl('https://my.feishu.cn/drive/folder/VOo7fZqi0lNLSvdQw2XcE0AqnUh')).toBe('VOo7fZqi0lNLSvdQw2XcE0AqnUh');
    expect(parseDriveFolderUrl('https://my.feishu.cn/drive/folder/YyGZfIOUjlpVO1dr8EPcorcun3g')).toBe('YyGZfIOUjlpVO1dr8EPcorcun3g');
    expect(() => parseDriveFolderUrl('https://example.invalid/drive/folder/x')).toThrow('example.invalid');
  });
  it('accepts the configured corporate document host for ordinary Drive folders', () => {
    expect(parseDriveFolderUrl('https://cibt.co.cibfintech.com/drive/folder/CorpFolderToken'))
      .toBe('CorpFolderToken');
  });

  it('maps the supplied version owner to the template spelling 版本负责任', async () => {
    const { service } = fixture({ ownerTypo: true });
    const manifest = await service.prepare({ ...input, fields: { ...input.fields, 版本负责人: 'Topson' } });
    expect(manifest.files[0].patches[0].after).toContain('Topson');
    await expect(service.prepare({ ...input, fields: { ...input.fields, 版本负责人: 'Topson', 版本负责任: 'Other' } })).rejects.toThrow('不一致');
  });
  it('previews all three template types with paginated Drive listing and no writes', async () => {
    const { service, calls } = fixture();
    const manifest = await service.prepare(input);
    expect(manifest.files.map(f => [f.type, f.patches.length])).toEqual([['docx', 1], ['sheet', 1], ['bitable', 1]]);
    expect(manifest.ignoredFolders).toEqual(['子目录']);
    expect(manifest.files[0].patches[0].after).toContain('张三');
    expect(calls.some(c => c.method !== 'GET')).toBe(false);
    expect(calls.filter(c => c.url === '/open-apis/drive/v1/files' && c.params?.folder_token === 'template')).toHaveLength(2);
  });

  it('accepts supplied fields absent from the template and reports them as unused', async () => {
    const { service, calls } = fixture();
    const manifest = await service.prepare({ ...input, fields: { ...input.fields, 下发日期: '20261030' } });
    expect(manifest.unusedFields).toEqual(['下发日期']);
    expect(calls.every(call => !JSON.stringify(call).includes('下发日期'))).toBe(true);
  });

  it('stops before approval when the request group can only view the destination folder', async () => {
    const { service, calls } = fixture({ parentGroupPerm: 'view' });
    await expect(service.prepare(input, caller)).rejects.toThrow('仅有查看权限');
    expect(calls.some(c => c.method !== 'GET')).toBe(false);
  });

  it('creates one folder, copies each file, replaces only copies, and rejects a replay', async () => {
    const { service, calls } = fixture();
    const manifest = await service.prepare(input);
    const jobId = await approvedJob(service, manifest);
    const result = await service.executeJob(jobId);
    expect(result.files).toHaveLength(3);
    expect(calls.filter(c => c.url.endsWith('/create_folder'))).toHaveLength(1);
    expect(calls.filter(c => c.url.endsWith('/copy'))).toHaveLength(3);
    expect(calls.filter(c => ['PATCH', 'PUT'].includes(c.method)).every(c => !/doxcnSource|shtSource|baseSource/.test(c.url))).toBe(true);
    await expect(service.executeJob(jobId)).rejects.toThrow('已执行或状态不明');
    const row = (await (await stateDb()).execute({ sql: 'SELECT status FROM release_jobs WHERE id = ?', args: [jobId] })).rows[0];
    expect(row.status).toBe('completed');
  });

  it('copies files without matching placeholders unchanged and skips content writes', async () => {
    const { service, calls } = fixture({ plainTemplates: true });
    const manifest = await service.prepare(input);
    expect(manifest.files.map(file => file.patches.length)).toEqual([0, 0, 0]);
    const result = await service.executeJob(await approvedJob(service, manifest));
    expect(result.files.map(file => file.name)).toEqual(['说明.docx', '清单', '责任表']);
    expect(calls.filter(call => call.url.endsWith('/copy'))).toHaveLength(3);
    expect(calls.filter(call => ['PATCH', 'PUT'].includes(call.method))).toHaveLength(0);
  });

  it('updates a Docx page title through Drive while keeping it out of the Docx body batch', async () => {
    const { service, calls } = fixture({ docTitleMarker: true });
    const manifest = await service.prepare(input);
    expect(manifest.files[0].patches).toHaveLength(2);
    const jobId = await approvedJob(service, manifest);
    const result = await service.executeJob(jobId);
    expect(result.files[0].name).toBe('说明20261029');
    const rename = calls.find(c => c.url === '/open-apis/drive/v1/files/doxcnCopy' && c.method === 'PATCH');
    expect(rename).toMatchObject({ params: { type: 'docx' }, data: { new_title: '说明20261029' } });
    const batch = calls.find(c => c.url.endsWith('/blocks/batch_update'));
    expect((batch?.data as { requests: Array<{ block_id: string }> }).requests.map(r => r.block_id)).toEqual(['blockCopy']);
  });

  it('previews and renames a copied Sheets file title without changing the template', async () => {
    const { service, calls } = fixture({ sheetTitleMarker: true });
    const manifest = await service.prepare(input);
    expect(manifest.files[1].patches[0]).toMatchObject({ location: '文件标题', before: '清单{{下发编号}}', after: '清单20261029', fileName: true });
    const jobId = await approvedJob(service, manifest);
    const result = await service.executeJob(jobId);
    expect(result.files[1].name).toBe('清单20261029');
    expect(calls.find(c => c.url === '/open-apis/drive/v1/files/shtCopy' && c.method === 'PATCH'))
      .toMatchObject({ params: { type: 'sheet' }, data: { new_title: '清单20261029' } });
    expect(calls.some(c => c.url === '/open-apis/drive/v1/files/shtSource' && c.method === 'PATCH')).toBe(false);
    const writes = calls.filter(c => c.url.includes('/sheets/v2/') && c.method === 'PUT');
    expect(writes).toHaveLength(1);
  });

  it('does not write a copied merged cell if its merge changes after scanning', async () => {
    const { service, calls } = fixture({ mergedSheet: true, copyMergeChangesAfterScan: true });
    const manifest = await service.prepare(input);
    expect(manifest.files[1].patches[0]).toMatchObject({ mergeRange: 'A1:B1', cell: 'A1' });
    const jobId = await approvedJob(service, manifest);
    await expect(service.executeJob(jobId)).rejects.toThrow('合并区域已变化');
    expect(calls.filter(call => call.url.includes('/sheets/v2/') && call.method === 'PUT')).toHaveLength(0);
  });

  it('writes and verifies a matched placeholder in the copied merged-cell anchor', async () => {
    const { service, calls } = fixture({ mergedSheet: true });
    const manifest = await service.prepare(input);
    expect(manifest.files[1].patches[0]).toMatchObject({ cell: 'A1', mergeRange: 'A1:B1' });
    await service.executeJob(await approvedJob(service, manifest));
    const writes = calls.filter(call => call.url.includes('/sheets/v2/') && call.method === 'PUT');
    expect(writes).toHaveLength(1);
    expect(writes[0].data).toMatchObject({ valueRange: { range: 'sheetCopy!A1:A1', values: [['2026-10-29']] } });
  });

  it('writes a mixed-text spreadsheet placeholder only in the approved copy', async () => {
    const { service, calls } = fixture({ sheetMixedText: true });
    const manifest = await service.prepare(input);
    expect(manifest.files[1].patches[0]).toMatchObject({ before: 'L_REL_{{下发编号}}', after: 'L_REL_20261029' });
    await service.executeJob(await approvedJob(service, manifest));
    const writes = calls.filter(call => call.url.includes('/sheets/v2/') && call.method === 'PUT');
    expect(writes).toHaveLength(1);
    expect(writes[0]).toMatchObject({ url: '/open-apis/sheets/v2/spreadsheets/shtCopy/values', data: { valueRange: { values: [['L_REL_20261029']] } } });
  });

  it('refuses to rename when the copied file title has changed since approval', async () => {
    const { service, calls } = fixture({ sheetTitleMarker: true });
    await expect(service.renameFile('shtCopy', 'sheet', 'newFolder', 'different', '清单20261029')).rejects.toThrow('已变化');
    expect(calls.some(c => c.method === 'PATCH')).toBe(false);
  });

  it('stops before writes for unsupported types, missing values, duplicate folder and changed source', async () => {
    await expect(fixture({ unsupported: true }).service.prepare(input)).rejects.toThrow('不支持的类型');
    await expect(fixture().service.prepare({ ...input, fields: { 日期: '2026-10-29' } })).rejects.toThrow('负责人');
    await expect(fixture({ duplicateFolder: true }).service.prepare(input)).rejects.toThrow('已有同名');
    const { service, calls } = fixture({ sourceChanged: true });
    const manifest = await service.prepare(input);
    const jobId = await approvedJob(service, manifest);
    await expect(service.executeJob(jobId)).rejects.toThrow('已变化');
    expect(calls.some(c => c.method !== 'GET')).toBe(false);
  });

  it('records uncertain copies without retrying them', async () => {
    const { service, calls } = fixture({ failCopy: true });
    const jobId = await approvedJob(service, await service.prepare(input));
    await expect(service.executeJob(jobId)).rejects.toThrow('copy timeout');
    expect(calls.filter(c => c.url.includes('shtSource/copy'))).toHaveLength(1);
    const row = (await (await stateDb()).execute({ sql: 'SELECT status, folder_token FROM release_jobs WHERE id = ?', args: [jobId] })).rows[0];
    expect(row.status).toBe('needs_review');
    expect(row.folder_token).toBe('newFolder');
  });

  it('records a forbidden folder creation as a known failure, without a copy', async () => {
    const { service, calls } = fixture({ failFolderForbidden: true });
    const jobId = await approvedJob(service, await service.prepare(input));
    await expect(service.executeJob(jobId)).rejects.toThrow('1061004');
    const row = (await (await stateDb()).execute({ sql: 'SELECT status, folder_token FROM release_jobs WHERE id = ?', args: [jobId] })).rows[0];
    expect(row.status).toBe('failed');
    expect(row.folder_token).toBeNull();
    expect(calls.filter(c => c.url.endsWith('/copy'))).toHaveLength(0);
  });

  it('does not approve a formula marker and waits for a copied document to become readable', async () => {
    await expect(fixture({ formulaMarker: true }).service.prepare(input)).rejects.toThrow('公式');
    const { service, calls } = fixture({ copyReadDelay: true });
    const jobId = await approvedJob(service, await service.prepare(input));
    await service.executeJob(jobId);
    expect(calls.filter(c => c.url.includes('doxcnCopy/blocks') && c.method === 'GET').length).toBeGreaterThan(1);
    expect(calls.filter(c => c.url.includes('doxcnSource/blocks/batch_update'))).toHaveLength(0);
  });

  it('marks an uncertain write for manual review and never retries it', async () => {
    const { service, calls } = fixture({ failWrite: true });
    const jobId = await approvedJob(service, await service.prepare(input));
    await expect(service.executeJob(jobId)).rejects.toThrow('write timeout');
    expect(calls.filter(c => c.url.endsWith('/blocks/batch_update'))).toHaveLength(1);
    const row = (await (await stateDb()).execute({ sql: 'SELECT status FROM release_files WHERE job_id = ? AND source_token = ?', args: [jobId, 'doxcnSource'] })).rows[0];
    expect(row.status).toBe('uncertain');
  });

  it('deduplicates a repeated preparation and invalidates an interrupted run', async () => {
    const { service } = fixture();
    const manifest = await service.prepare(input);
    const first = await service.createJob(caller, manifest, 1234567890);
    const second = await service.createJob(caller, manifest, 1234567890);
    expect(second).toMatchObject({ jobId: first.jobId, reused: true });
    const db = await stateDb();
    await db.execute({ sql: "UPDATE release_jobs SET status = 'running' WHERE id = ?", args: [first.jobId] });
    const { recoverInterruptedWorkflows } = await import('../state/store');
    await recoverInterruptedWorkflows();
    const row = (await db.execute({ sql: 'SELECT status FROM release_jobs WHERE id = ?', args: [first.jobId] })).rows[0];
    expect(row.status).toBe('needs_review');
    await expect(service.executeJob(first.jobId)).rejects.toThrow('已执行或状态不明');
  });

  it('reports a failed approval as failed when the same message is deduplicated', async () => {
    const { service } = fixture();
    const manifest = await service.prepare(input);
    const first = await service.createJob(caller, manifest, 1234567891);
    const db = await stateDb();
    await db.execute({ sql: "UPDATE release_jobs SET status = 'failed' WHERE id = ?", args: [first.jobId] });
    const repeated = await service.createJob(caller, manifest, 1234567891);
    expect(repeated).toMatchObject({ jobId: first.jobId, reused: true, status: 'failed' });
  });

  it('refuses a Studio workflow run without a confirmed approval', async () => {
    const { service, calls } = fixture();
    const { jobId } = await service.createJob(caller, await service.prepare(input));
    await expect(service.executeJob(jobId)).rejects.toThrow('缺少原请求者');
    expect(calls.some(c => c.method !== 'GET')).toBe(false);
  });
});
