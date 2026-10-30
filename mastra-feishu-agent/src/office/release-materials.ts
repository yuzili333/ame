import { createHash, randomUUID } from 'node:crypto';
import { z } from 'zod';
import { config } from '../config';
import type { Caller } from '../identity';
import { isMessageRecalled, stateDb } from '../state/store';
import type { OfficeApi } from './client';

const token = z.string().min(1).max(256).regex(/^[\w@.-]+$/);
const fieldName = z.string().trim().min(1).max(100);
export const releaseInput = z.object({
  templateFolderUrl: z.string().url(),
  materialsParentUrl: z.string().url(),
  releaseId: z.string().trim().min(1).max(100).regex(/^[^\\/<>:"|?*\x00-\x1f]+$/),
  fields: z.record(fieldName, z.string().max(20_000).refine(v => !v.includes('{{') && !v.includes('}}'), '字段值不能包含占位符标记')),
  sheetScopes: z.array(z.object({ fileName: z.string().min(1), sheetName: z.string().min(1), range: z.string().regex(/^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$/) })).max(100).optional(),
});
export type ReleaseInput = z.infer<typeof releaseInput>;
export type ReleaseType = 'docx' | 'sheet' | 'bitable';
export type DriveFile = { token: string; name: string; type: ReleaseType | 'folder' | string; url?: string; modified_time?: string };
type Patch = { location: string; field: string; before: string; after: string; blockId?: string; elements?: unknown[]; title?: boolean; fileName?: boolean; sheetId?: string; cell?: string; mergeRange?: string; tableId?: string; recordId?: string; fieldName?: string };
type ScannedFile = DriveFile & { patches: Patch[]; fingerprint: string };
type ScanResult = { patches: Patch[]; fingerprint: string };
export type ReleaseManifest = { input: ReleaseInput; templateFolderToken: string; materialsParentToken: string; files: ScannedFile[]; ignoredFolders: string[]; unusedFields?: string[] };
const marker = /\{\{([^{}]+)\}\}/g;
const enc = encodeURIComponent;

export function parseDriveFolderUrl(value: string): string {
  const url = new URL(value);
  if (url.protocol !== 'https:') throw new Error('请提供 HTTPS 飞书文件夹链接');
  if (!config.documentHosts.includes(url.hostname.toLowerCase())) throw new Error(`文件夹链接域名 ${url.hostname} 不在 FEISHU_DOCUMENT_HOSTS 中；请将客户端分享链接的域名加入配置`);
  const match = /^\/drive\/folder\/([\w@.-]+)\/?$/.exec(url.pathname);
  if (!match) throw new Error('仅支持普通云空间 /drive/folder/ 文件夹链接，不支持 Wiki 目录');
  return token.parse(match[1]);
}

function replacement(before: string, fields: Record<string, string>, used: Set<string>, preserveUnknown = false): string {
  return before.replace(marker, (whole, raw: string) => {
    const name = raw.trim();
    if (!Object.hasOwn(fields, name)) {
      if (preserveUnknown) return whole;
      throw new Error(`模板需要字段“${name}”，请补充后重新发起`);
    }
    used.add(name);
    return fields[name];
  });
}
function containsMarker(value: unknown): boolean { return (JSON.stringify(value) ?? '').includes('{{'); }
function same(a: unknown, b: unknown): boolean { return JSON.stringify(a) === JSON.stringify(b); }
function column(index: number): string { let out = ''; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return out; }
function columnIndex(label: string): number { return [...label].reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1; }
type SheetScope = NonNullable<ReleaseInput['sheetScopes']>[number];
type Merge = { start_row_index?: number; end_row_index?: number; start_column_index?: number; end_column_index?: number };
type Sheet = { sheet_id?: string; title?: string; grid_properties?: { row_count?: number; column_count?: number }; merges?: Merge[] };
function mergeAt(sheet: Sheet, row: number, col: number): { anchor: string; range: string } | undefined {
  for (const merge of sheet.merges ?? []) {
    const { start_row_index: r0, end_row_index: r1, start_column_index: c0, end_column_index: c1 } = merge;
    // Sheets query reports both end indexes inclusively. A one-cell range
    // therefore has identical start and end indexes.
    if (![r0, r1, c0, c1].every(n => Number.isInteger(n)) || r0! < 0 || c0! < 0 || r1! < r0! || c1! < c0!) throw new Error(`${sheet.title ?? sheet.sheet_id} 的合并区域元数据无效`);
    if (row >= r0! && row <= r1! && col >= c0! && col <= c1!) {
      const anchor = `${column(c0!)}${r0! + 1}`;
      return { anchor, range: `${anchor}:${column(c1!)}${r1! + 1}` };
    }
  }
  return undefined;
}

export class ReleaseMaterialsService {
  constructor(private readonly api: OfficeApi) {}

  async listFolder(folderToken: string): Promise<DriveFile[]> {
    const files: DriveFile[] = [];
    let pageToken: string | undefined;
    do {
      const data = await this.api.call({ method: 'GET', url: '/open-apis/drive/v1/files', params: { folder_token: token.parse(folderToken), page_size: 200, ...(pageToken ? { page_token: pageToken } : {}) } });
      const page = (data.files as Array<Record<string, unknown>> | undefined) ?? [];
      for (const item of page) {
        const itemToken = item.token ?? item.file_token;
        if (typeof itemToken !== 'string' || typeof item.name !== 'string' || typeof item.type !== 'string') throw new Error('飞书目录返回的文件元数据不完整');
        files.push({ token: itemToken, name: item.name, type: item.type, url: typeof item.url === 'string' ? item.url : undefined, modified_time: typeof item.modified_time === 'string' ? item.modified_time : undefined });
      }
      if (files.length > 1000) throw new Error('当前层文件超过 1000 个，无法安全预览整批操作');
      pageToken = data.has_more === true ? String(data.next_page_token ?? data.page_token ?? '') : undefined;
      if (data.has_more === true && !pageToken) throw new Error('飞书目录分页缺少下一页标识');
    } while (pageToken);
    return files;
  }

  async createFolder(parentToken: string, name: string): Promise<{ token: string; url?: string }> {
    const data = await this.api.call({ method: 'POST', url: '/open-apis/drive/v1/files/create_folder', data: { folder_token: token.parse(parentToken), name } });
    if (typeof data.token !== 'string') throw new Error('创建文件夹结果不明：响应缺少 token，请人工核对目标目录');
    return { token: data.token, url: typeof data.url === 'string' ? data.url : undefined };
  }

  async copyFile(file: Pick<DriveFile, 'token' | 'name' | 'type'>, folderToken: string): Promise<{ token: string; url?: string }> {
    if (!['docx', 'sheet', 'bitable'].includes(file.type)) throw new Error(`不支持复制 ${file.type} 文件`);
    const data = await this.api.call({ method: 'POST', url: `/open-apis/drive/v1/files/${enc(token.parse(file.token))}/copy`, data: { name: file.name, type: file.type, folder_token: token.parse(folderToken) } });
    const copied = data.file as Record<string, unknown> | undefined;
    const copyToken = copied?.token ?? copied?.file_token;
    if (typeof copyToken !== 'string') throw new Error(`复制 ${file.name} 的结果不明：响应缺少副本 token，请人工核对`);
    if (copyToken === file.token) throw new Error(`复制 ${file.name} 返回了源文件 token，请人工核对，不得修改模板`);
    return { token: copyToken, url: typeof copied?.url === 'string' ? copied.url : undefined };
  }

  async renameFile(fileToken: string, fileType: ReleaseType, folderToken: string, expectedName: string, newName: string): Promise<{ name: string }> {
    let current: DriveFile | undefined;
    for (let attempt = 0; attempt < 5; attempt++) {
      current = (await this.listFolder(folderToken)).find(file => file.token === fileToken);
      if (current) break;
      if (attempt < 4) await new Promise(resolve => setTimeout(resolve, 300 * (attempt + 1)));
    }
    if (!current || current.type !== fileType || current.name !== expectedName) {
      throw new Error('文件标题或所在目录已变化，未执行改名；请重新预览');
    }
    await this.api.call({ method: 'PATCH', url: `/open-apis/drive/v1/files/${enc(token.parse(fileToken))}`, params: { type: fileType }, data: { new_title: newName } });
    for (let attempt = 0; attempt < 5; attempt++) {
      const updated = (await this.listFolder(folderToken)).find(file => file.token === fileToken);
      if (updated?.name === newName) return { name: newName };
      if (updated && updated.name !== expectedName) break;
      if (attempt < 4) await new Promise(resolve => setTimeout(resolve, 300 * (attempt + 1)));
    }
    throw new Error('文件改名后的回读结果不符，请人工核对，不得自动重试');
  }

  private async scanDocx(file: DriveFile, fields: Record<string, string>, used: Set<string>): Promise<ScanResult> {
    const patches: Patch[] = [];
    const hash = createHash('sha256');
    let pageToken: string | undefined;
    let index = 0;
    do {
      const data = await this.api.call({ method: 'GET', url: `/open-apis/docx/v1/documents/${enc(file.token)}/blocks`, params: { page_size: 500, ...(pageToken ? { page_token: pageToken } : {}) } });
      const blocks = (data.items as Array<Record<string, unknown>> | undefined) ?? [];
      hash.update(JSON.stringify(blocks));
      for (const block of blocks) {
        const location = `块 ${++index}`;
        if (!containsMarker(block)) continue;
        const candidates = Object.entries(block).filter(([, v]) => v && typeof v === 'object' && Array.isArray((v as { elements?: unknown }).elements)) as Array<[string, { elements: Array<Record<string, unknown>> }]>;
        if (candidates.length !== 1 || typeof block.block_id !== 'string') throw new Error(`${file.name} ${location} 的占位符位于不支持的块格式`);
        const [contentKey, content] = candidates[0];
        const elements = structuredClone(content.elements);
        let changed = false;
        for (const element of elements) {
          const run = element.text_run as { content?: string } | undefined;
          if (typeof run?.content !== 'string') {
            if (containsMarker(element)) throw new Error(`${file.name} ${location} 的占位符跨越富文本元素或位于非文本元素`);
            continue;
          }
          const next = replacement(run.content, fields, used);
          changed ||= next !== run.content;
          run.content = next;
        }
        const combined = elements.map(element => (element.text_run as { content?: string } | undefined)?.content ?? '').join('');
        if (!changed || combined.includes('{{') || containsMarker({ ...block, [contentKey]: { ...content, elements } })) throw new Error(`${file.name} ${location} 的占位符跨越富文本元素或位于不支持的属性`);
        const title = block.block_type === 1;
        if (title && (contentKey !== 'page' || elements.some(element => typeof (element.text_run as { content?: unknown } | undefined)?.content !== 'string'))) {
          throw new Error(`${file.name} 的标题包含不支持的富文本格式，无法安全替换`);
        }
        patches.push({ location, field: title ? '标题' : '富文本', before: JSON.stringify(content.elements), after: JSON.stringify(elements), blockId: block.block_id, elements, ...(title ? { title: true } : {}) });
      }
      pageToken = data.has_more === true ? String(data.page_token ?? '') : undefined;
      if (data.has_more === true && !pageToken) throw new Error('Docx 分页缺少下一页标识');
    } while (pageToken);
    return { patches, fingerprint: hash.digest('hex') };
  }

  private async scanSheet(file: DriveFile, fields: Record<string, string>, used: Set<string>, scopes: SheetScope[] = []): Promise<ScanResult> {
    const data = await this.api.call({ method: 'GET', url: `/open-apis/sheets/v3/spreadsheets/${enc(file.token)}/sheets/query` });
    const hash = createHash('sha256');
    const sheets = (data.sheets as Sheet[] | undefined) ?? [];
    hash.update(JSON.stringify(sheets));
    const patches: Patch[] = [];
    const fileScopes = scopes.filter(scope => scope.fileName === file.name);
    const matchedScopes = new Set<string>();
    for (const sheet of sheets) {
      if (!sheet.sheet_id) throw new Error(`${file.name} 有工作表缺少 sheet_id`);
      const scope = fileScopes.find(item => item.sheetName === sheet.title);
      if (fileScopes.length && !scope) continue;
      if (scope) {
        if (matchedScopes.has(scope.sheetName) || fileScopes.filter(item => item.sheetName === scope.sheetName).length !== 1) throw new Error(`${file.name} 的 ${scope.sheetName} 扫描范围重复`);
        matchedScopes.add(scope.sheetName);
      }
      const rows = sheet.grid_properties?.row_count;
      const cols = sheet.grid_properties?.column_count;
      if (rows === undefined || cols === undefined) throw new Error(`${file.name} 的工作表缺少行列范围，无法安全完整扫描`);
      if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 0 || cols < 0) throw new Error(`${file.name} 的工作表范围无效，无法安全扫描`);
      if (!rows || !cols) {
        if (scope) throw new Error(`${file.name} ${sheet.title} 是空工作表，不能使用扫描范围 ${scope.range}`);
        continue;
      }
      let firstRow = 1, lastRow = rows, firstCol = 0, lastCol = cols - 1;
      if (scope) {
        const parts = /^([A-Z]+)([1-9]\d*):([A-Z]+)([1-9]\d*)$/.exec(scope.range)!;
        firstCol = columnIndex(parts[1]); firstRow = Number(parts[2]); lastCol = columnIndex(parts[3]); lastRow = Number(parts[4]);
        if (firstCol > lastCol || firstRow > lastRow || lastCol >= cols || lastRow > rows) throw new Error(`${file.name} ${sheet.title} 扫描范围 ${scope.range} 超出工作表`);
      } else if (rows * cols > 100_000) {
        throw new Error(`${file.name} / ${sheet.title ?? sheet.sheet_id} 扫描范围过大（${rows}×${cols}）；请指定 sheetScopes 中的 fileName、sheetName 和 range`);
      }
      const width = lastCol - firstCol + 1;
      if (width > 1000 || (lastRow - firstRow + 1) * width > 100_000) throw new Error(`${file.name} / ${sheet.title ?? sheet.sheet_id} 扫描范围过大；请缩小 sheetScopes 的 range`);
      const rowsPerRead = Math.max(1, Math.floor(1000 / width));
      for (let start = firstRow; start <= lastRow; start += rowsPerRead) {
        const end = Math.min(start + rowsPerRead - 1, lastRow);
        const range = `${token.parse(sheet.sheet_id)}!${column(firstCol)}${start}:${column(lastCol)}${end}`;
        let values: Record<string, unknown>;
        try {
          values = await this.api.call({ method: 'GET', url: `/open-apis/sheets/v2/spreadsheets/${enc(file.token)}/values/${range}`, params: { valueRenderOption: 'Formula' } });
        } catch (error) {
          throw new Error(`${file.name} / ${sheet.title ?? sheet.sheet_id} / ${range} 读取失败：${String(error)}`);
        }
        const matrix = (values.valueRange as { values?: unknown[][] } | undefined)?.values ?? [];
        hash.update(JSON.stringify(matrix));
        for (let r = 0; r < matrix.length; r++) for (let c = 0; c < matrix[r].length; c++) {
          const cell = matrix[r][c];
          if (!containsMarker(cell)) continue;
          const matchedField = [...(JSON.stringify(cell) ?? '').matchAll(marker)]
            .some(([, raw]) => Object.hasOwn(fields, raw.trim()));
          if (!matchedField) continue;
          const actualCol = firstCol + c;
          if (typeof cell !== 'string' || cell.trimStart().startsWith('=')) throw new Error(`${file.name} ${sheet.title ?? sheet.sheet_id}!${column(actualCol)}${start + r} 含公式或非普通文本占位符`);
          const cellRef = `${column(actualCol)}${start + r}`;
          const merge = mergeAt(sheet, start + r - 1, actualCol);
          if (merge && merge.anchor !== cellRef) throw new Error(`${file.name} ${sheet.title ?? sheet.sheet_id}!${cellRef} 位于合并区域 ${merge.range} 的非左上角，无法安全替换`);
          const after = replacement(cell, fields, used, true);
          const matchedNames = [...new Set([...cell.matchAll(marker)].map(([, raw]) => raw.trim()).filter(name => Object.hasOwn(fields, name)))];
          patches.push({ location: `${sheet.title ?? sheet.sheet_id}!${merge?.range ?? cellRef}`, field: matchedNames.join('、'), before: cell, after, sheetId: sheet.sheet_id, cell: cellRef, ...(merge ? { mergeRange: merge.range } : {}) });
        }
      }
    }
    if (matchedScopes.size !== fileScopes.length) throw new Error(`${file.name} 的 sheetScopes 含不存在的工作表`);
    return { patches, fingerprint: hash.digest('hex') };
  }

  private async scanBitable(file: DriveFile, fields: Record<string, string>, used: Set<string>): Promise<ScanResult> {
    const patches: Patch[] = [];
    const hash = createHash('sha256');
    let tablePage: string | undefined;
    do {
      const tables = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(file.token)}/tables`, params: { page_size: 100, ...(tablePage ? { page_token: tablePage } : {}) } });
      hash.update(JSON.stringify(tables.items ?? []));
      for (const table of (tables.items as Array<{ table_id?: string; name?: string }> | undefined) ?? []) {
        if (!table.table_id) throw new Error(`${file.name} 有数据表缺少 table_id`);
        const types = new Map<string, number>();
        let fieldPage: string | undefined;
        do {
          const fieldData = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(file.token)}/tables/${enc(table.table_id)}/fields`, params: { page_size: 100, ...(fieldPage ? { page_token: fieldPage } : {}) } });
          hash.update(JSON.stringify(fieldData.items ?? []));
          for (const field of (fieldData.items as Array<{ field_name?: string; type?: number }> | undefined) ?? []) if (field.field_name) types.set(field.field_name, field.type ?? -1);
          fieldPage = fieldData.has_more ? String(fieldData.page_token ?? '') : undefined;
          if (fieldData.has_more && !fieldPage) throw new Error('Bitable 字段分页缺少标识');
        } while (fieldPage);
        let recordPage: string | undefined;
        do {
          const records = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(file.token)}/tables/${enc(table.table_id)}/records`, params: { page_size: 100, ...(recordPage ? { page_token: recordPage } : {}) } });
          hash.update(JSON.stringify(records.items ?? []));
          for (const record of (records.items as Array<{ record_id?: string; fields?: Record<string, unknown> }> | undefined) ?? []) {
            if (!record.record_id) throw new Error(`${file.name} 有记录缺少 record_id`);
            for (const [name, value] of Object.entries(record.fields ?? {})) {
              if (!containsMarker(value)) continue;
              if (types.get(name) !== 1) throw new Error(`${file.name} / ${table.name ?? table.table_id} / ${name} 的占位符不在普通文本字段`);
              if (Array.isArray(value) && value.some(v => typeof v?.text !== 'string')) throw new Error(`${file.name} / ${name} 包含复杂文本元素，无法安全替换`);
              const text = typeof value === 'string' ? value : Array.isArray(value) ? value.map(v => v.text as string).join('') : '';
              if (!/^\{\{[^{}]+\}\}$/.test(text)) throw new Error(`${file.name} / ${table.name ?? table.table_id} / ${name} 不是整值占位符`);
              patches.push({ location: `${table.name ?? table.table_id} / ${record.record_id} / ${name}`, field: text.slice(2, -2).trim(), before: text, after: replacement(text, fields, used), tableId: table.table_id, recordId: record.record_id, fieldName: name });
            }
          }
          recordPage = records.has_more ? String(records.page_token ?? '') : undefined;
          if (records.has_more && !recordPage) throw new Error('Bitable 记录分页缺少标识');
        } while (recordPage);
      }
      tablePage = tables.has_more ? String(tables.page_token ?? '') : undefined;
      if (tables.has_more && !tablePage) throw new Error('Bitable 数据表分页缺少标识');
    } while (tablePage);
    return { patches, fingerprint: hash.digest('hex') };
  }

  async scanFile(file: DriveFile, fields: Record<string, string>, used = new Set<string>(), sheetScopes: SheetScope[] = []): Promise<ScannedFile> {
    let scan: ScanResult;
    try {
      scan = file.type === 'docx' ? await this.scanDocx(file, fields, used)
        : file.type === 'sheet' ? await this.scanSheet(file, fields, used, sheetScopes)
        : file.type === 'bitable' ? await this.scanBitable(file, fields, used)
        : (() => { throw new Error(`不支持的模板类型：${file.name} (${file.type})`); })();
    } catch (error) {
      throw new Error(`扫描模板“${file.name}”（${file.type}）失败：${String(error)}`, { cause: error });
    }
    if ((file.type === 'sheet' || file.type === 'bitable') && file.name.includes('{{')) {
      const after = replacement(file.name, fields, used);
      if (after === file.name || after.includes('{{') || after.length > 255) throw new Error(`${file.name} 的文件标题占位符无法安全替换`);
      scan.patches.unshift({ location: '文件标题', field: '标题', before: file.name, after, fileName: true });
    }
    return { ...file, ...scan };
  }

  async prepare(input: ReleaseInput, caller?: Caller): Promise<ReleaseManifest> {
    const parsed = releaseInput.parse(input);
    if (parsed.fields['下发编号'] !== undefined && parsed.fields['下发编号'] !== parsed.releaseId) throw new Error('“下发编号”字段值必须与文件夹下发编号一致');
    const owner = parsed.fields['版本负责人'];
    const ownerTypo = parsed.fields['版本负责任'];
    if (owner !== undefined && ownerTypo !== undefined && owner !== ownerTypo) throw new Error('“版本负责人”和“版本负责任”的值不一致，请明确提供同一个负责人');
    if (owner !== undefined || ownerTypo !== undefined) {
      parsed.fields = { ...parsed.fields, 版本负责人: owner ?? ownerTypo!, 版本负责任: owner ?? ownerTypo! };
    }
    parsed.fields = { 下发编号: parsed.releaseId, ...parsed.fields };
    const templateFolderToken = parseDriveFolderUrl(parsed.templateFolderUrl);
    const materialsParentToken = parseDriveFolderUrl(parsed.materialsParentUrl);
    if (templateFolderToken === materialsParentToken) throw new Error('模板目录和材料父目录不能相同');
    const [source, destination] = await Promise.all([this.listFolder(templateFolderToken), this.listFolder(materialsParentToken)]);
    if (destination.some(f => f.name === parsed.releaseId)) throw new Error(`材料父目录已有同名项目“${parsed.releaseId}”，未创建`);
    if (caller?.chatType === 'group') {
      const permissions = await this.api.call({ method: 'GET', url: `/open-apis/drive/v1/permissions/${enc(materialsParentToken)}/members`, params: { type: 'folder' } });
      const members = (permissions.items as Array<{ member_type?: string; member_id?: string; perm?: string }> | undefined) ?? [];
      const currentGroup = members.filter(member => member.member_type === 'openchat' && member.member_id === caller.chatId);
      if (currentGroup.length && currentGroup.every(member => member.perm === 'view')) {
        throw new Error('材料父目录对当前群仅有查看权限；请将该群在材料父目录的权限设为可编辑或可管理后重新发起。尚未发送写入审批卡片。');
      }
    }
    const ignoredFolders = source.filter(f => f.type === 'folder').map(f => f.name);
    const files = source.filter(f => f.type !== 'folder');
    if (!files.length) throw new Error('模板文件夹当前层没有可复制的文档');
    const unsupported = files.filter(f => !['docx', 'sheet', 'bitable'].includes(f.type));
    if (unsupported.length) throw new Error(`模板文件夹含不支持的类型：${unsupported.map(f => `${f.name} (${f.type})`).join('、')}`);
    for (const scope of parsed.sheetScopes ?? []) if (!files.some(file => file.type === 'sheet' && file.name === scope.fileName)) throw new Error(`sheetScopes 指定的表格文件不存在：${scope.fileName}`);
    const used = new Set<string>();
    const scanned: ScannedFile[] = [];
    for (const file of files) scanned.push(await this.scanFile(file, parsed.fields, used, parsed.sheetScopes));
    const unusedFields = Object.keys(parsed.fields).filter(name => name !== '下发编号' && !used.has(name)
      && !(name === '版本负责人' && used.has('版本负责任'))
      && !(name === '版本负责任' && used.has('版本负责人')));
    return { input: parsed, templateFolderToken, materialsParentToken, files: scanned, ignoredFolders, unusedFields };
  }

  async createJob(caller: Caller, manifest: ReleaseManifest, messageCreatedAt = Date.now()): Promise<{ jobId: string; reused: boolean; status: string; approvalId?: string }> {
    if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，本轮下发已取消');
    const db = await stateDb();
    const jobId = randomUUID();
    const requestKey = createHash('sha256').update(JSON.stringify([caller.userId, caller.chatId, messageCreatedAt, manifest.input])).digest('hex');
    const inserted = await db.execute({ sql: `INSERT OR IGNORE INTO release_jobs(id, requester_id, chat_id, status, manifest_json, request_key, created_at, updated_at, origin_message_id) VALUES (?, ?, ?, 'pending', ?, ?, ?, ?, ?)`, args: [jobId, caller.userId, caller.chatId, JSON.stringify(manifest), requestKey, Date.now(), Date.now(), caller.messageId ?? null] });
    if (await isMessageRecalled(caller.messageId)) {
      await db.execute({ sql: "UPDATE release_jobs SET status = 'denied', error = '原始消息已撤回' WHERE id = ? AND status = 'pending'", args: [jobId] });
      throw new Error('原始消息已撤回，本轮下发已取消');
    }
    if (!inserted.rowsAffected) {
      const existing = (await db.execute({ sql: 'SELECT id, status, approval_id FROM release_jobs WHERE request_key = ?', args: [requestKey] })).rows[0];
      return { jobId: String(existing.id), reused: true, status: String(existing.status), approvalId: typeof existing.approval_id === 'string' ? existing.approval_id : undefined };
    }
    for (const file of manifest.files) await db.execute({ sql: `INSERT INTO release_files(job_id, source_token, source_name, type, status) VALUES (?, ?, ?, ?, 'pending')`, args: [jobId, file.token, file.name, file.type] });
    return { jobId, reused: false, status: 'pending' };
  }

  async executeJob(jobId: string): Promise<{ folderUrl: string; files: Array<{ name: string; url: string }> }> {
    const db = await stateDb();
    const row = (await db.execute({ sql: 'SELECT * FROM release_jobs WHERE id = ?', args: [jobId] })).rows[0];
    if (!row || row.status !== 'pending') throw new Error('整批任务已执行或状态不明，不得重试');
    const approval = row.approval_id ? (await db.execute({ sql: 'SELECT status, requester_id, chat_id FROM actions WHERE id = ? AND kind = ?', args: [row.approval_id, 'release_materials'] })).rows[0] : undefined;
    if (!approval || approval.status !== 'executing' || approval.requester_id !== row.requester_id || approval.chat_id !== row.chat_id) throw new Error('整批任务缺少原请求者已确认的有效审批');
    const manifest = JSON.parse(String(row.manifest_json)) as ReleaseManifest;
    const claimed = await db.execute({ sql: "UPDATE release_jobs SET status = 'running', updated_at = ? WHERE id = ? AND status = 'pending'", args: [Date.now(), jobId] });
    if (!claimed.rowsAffected) throw new Error('整批任务已被其他执行者领取');
    try {
      const current = await this.prepare(manifest.input);
      if (!same(current, manifest)) throw new Error('审批后源文件、占位符或目标目录已变化，请重新预览并确认');
      await db.execute({ sql: "UPDATE release_jobs SET status = 'creating_folder', updated_at = ? WHERE id = ?", args: [Date.now(), jobId] });
      const folder = await this.createFolder(manifest.materialsParentToken, manifest.input.releaseId);
      await db.execute({ sql: "UPDATE release_jobs SET status = 'running', folder_token = ?, folder_url = ?, updated_at = ? WHERE id = ?", args: [folder.token, folder.url ?? null, Date.now(), jobId] });
      const resultFiles: Array<{ name: string; url: string }> = [];
      for (const file of manifest.files) {
        await db.execute({ sql: "UPDATE release_files SET status = 'copying' WHERE job_id = ? AND source_token = ?", args: [jobId, file.token] });
        const copy = await this.copyFile(file, folder.token);
        await db.execute({ sql: "UPDATE release_files SET status = 'waiting', copy_token = ?, copy_url = ? WHERE job_id = ? AND source_token = ?", args: [copy.token, copy.url ?? null, jobId, file.token] });
        if (!file.patches.length) {
          let visible = false;
          for (let attempt = 0; attempt < 8; attempt++) {
            const entry = (await this.listFolder(folder.token)).find(item => item.token === copy.token);
            if (entry) {
              if (entry.type !== file.type || entry.name !== file.name) throw new Error(`副本 ${file.name} 的类型或文件名与模板不一致，已停止`);
              visible = true;
              break;
            }
            if (attempt < 7) await new Promise(resolve => setTimeout(resolve, 500 * (attempt + 1)));
          }
          if (!visible) throw new Error(`副本 ${file.name} 尚未出现在材料目录中，请人工核对`);
          await db.execute({ sql: "UPDATE release_files SET status = 'completed' WHERE job_id = ? AND source_token = ?", args: [jobId, file.token] });
          resultFiles.push({ name: file.name, url: copy.url ?? `${new URL(manifest.input.materialsParentUrl).origin}/${file.type === 'sheet' ? 'sheets' : file.type === 'bitable' ? 'base' : 'docx'}/${copy.token}` });
          continue;
        }
        let scanned: ScannedFile | undefined;
        for (let attempt = 0; attempt < 8; attempt++) {
          try { scanned = await this.scanFile({ ...file, token: copy.token }, manifest.input.fields, new Set(), manifest.input.sheetScopes); break; }
          catch (error) { if (attempt === 7) throw new Error(`副本 ${file.name} 尚不可读或扫描失败：${String(error)}`); await new Promise(resolve => setTimeout(resolve, 500 * (attempt + 1))); }
        }
        if (!scanned || !same(scanned.patches.map(p => [p.location.replace(/rec[\w]+/, 'record'), p.field, p.before, p.after]), file.patches.map(p => [p.location.replace(/rec[\w]+/, 'record'), p.field, p.before, p.after]))) throw new Error(`副本 ${file.name} 与审批预览不一致，已停止`);
        await db.execute({ sql: "UPDATE release_files SET status = 'writing' WHERE job_id = ? AND source_token = ?", args: [jobId, file.token] });
        await this.writeCopy(scanned, folder.token);
        await this.verifyCopy(scanned, folder.token);
        await db.execute({ sql: "UPDATE release_files SET status = 'completed' WHERE job_id = ? AND source_token = ?", args: [jobId, file.token] });
        const titlePatch = scanned.patches.find(p => p.title);
        const fileNamePatch = scanned.patches.find(p => p.fileName);
        const name = titlePatch
          ? titlePatch.elements!.map(element => (element as { text_run: { content: string } }).text_run.content).join('')
          : fileNamePatch?.after ?? file.name;
        resultFiles.push({ name, url: copy.url ?? `${new URL(manifest.input.materialsParentUrl).origin}/${file.type === 'sheet' ? 'sheets' : file.type === 'bitable' ? 'base' : 'docx'}/${copy.token}` });
      }
      const folderUrl = folder.url ?? `${new URL(manifest.input.materialsParentUrl).origin}/drive/folder/${folder.token}`;
      await db.execute({ sql: "UPDATE release_jobs SET status = 'completed', result_json = ?, updated_at = ? WHERE id = ?", args: [JSON.stringify({ folderUrl, files: resultFiles }), Date.now(), jobId] });
      return { folderUrl, files: resultFiles };
    } catch (error) {
      const current = (await db.execute({ sql: 'SELECT status, folder_token FROM release_jobs WHERE id = ?', args: [jobId] })).rows[0];
      const knownCreateDenial = /^Error: Feishu API (1061004|99991672):/.test(String(error));
      const uncertain = !!current?.folder_token || (current?.status === 'creating_folder' && !knownCreateDenial);
      await db.execute({ sql: 'UPDATE release_jobs SET status = ?, error = ?, updated_at = ? WHERE id = ?', args: [uncertain ? 'needs_review' : 'failed', String(error), Date.now(), jobId] });
      await db.execute({ sql: "UPDATE release_files SET status = 'uncertain', error = ? WHERE job_id = ? AND status IN ('copying', 'writing')", args: [String(error), jobId] });
      throw error;
    }
  }

  private async writeCopy(file: ScannedFile, folderToken: string): Promise<void> {
    if (file.type === 'sheet') await this.assertSheetMerges(file);
    const fileNamePatch = file.patches.find(p => p.fileName);
    if (fileNamePatch) await this.renameFile(file.token, file.type as ReleaseType, folderToken, fileNamePatch.before, fileNamePatch.after);
    if (file.type === 'docx') {
      const titles = file.patches.filter(p => p.title);
      if (titles.length > 1) throw new Error(`${file.name} 有多个标题占位符，无法安全替换`);
      if (titles.length) {
        const title = titles[0].elements!.map(element => (element as { text_run: { content: string } }).text_run.content).join('');
        await this.api.call({ method: 'PATCH', url: `/open-apis/drive/v1/files/${enc(file.token)}`, params: { type: 'docx' }, data: { new_title: title } });
      }
      const body = file.patches.filter(p => !p.title);
      for (let i = 0; i < body.length; i += 50) {
        const batch = body.slice(i, i + 50);
        await this.api.call({ method: 'PATCH', url: `/open-apis/docx/v1/documents/${enc(file.token)}/blocks/batch_update`, data: { requests: batch.map(p => ({ block_id: p.blockId, update_text_elements: { elements: p.elements } })) } });
        if (i + 50 < body.length) await new Promise(resolve => setTimeout(resolve, 400));
      }
      return;
    }
    if (file.type === 'sheet') {
      for (const patch of file.patches.filter(p => !p.fileName)) await this.api.call({ method: 'PUT', url: `/open-apis/sheets/v2/spreadsheets/${enc(file.token)}/values`, data: { valueRange: { range: `${patch.sheetId}!${patch.cell}:${patch.cell}`, values: [[patch.after]] } } });
      return;
    }
    for (const patch of file.patches.filter(p => !p.fileName)) await this.api.call({ method: 'PUT', url: `/open-apis/bitable/v1/apps/${enc(file.token)}/tables/${enc(patch.tableId!)}/records/${enc(patch.recordId!)}`, data: { fields: { [patch.fieldName!]: patch.after } } });
  }

  private async verifyCopy(file: ScannedFile, folderToken: string): Promise<void> {
    if (file.type === 'sheet') await this.assertSheetMerges(file);
    for (const patch of file.patches) {
      if (patch.fileName) {
        const actual = (await this.listFolder(folderToken)).find(item => item.token === file.token);
        if (actual?.name !== patch.after) throw new Error(`${file.name} 文件标题回读内容不符，请人工核对`);
      } else if (file.type === 'docx') {
        const data = await this.api.call({ method: 'GET', url: `/open-apis/docx/v1/documents/${enc(file.token)}/blocks/${enc(patch.blockId!)}` });
        const block = data.block as Record<string, unknown> | undefined;
        const candidates = Object.values(block ?? {}).filter(v => v && typeof v === 'object' && Array.isArray((v as { elements?: unknown }).elements)) as Array<{ elements: unknown[] }>;
        const actual = candidates[0]?.elements;
        const titleMatches = patch.title && actual?.map(element => (element as { text_run?: { content?: string } }).text_run?.content ?? '').join('')
          === patch.elements?.map(element => (element as { text_run?: { content?: string } }).text_run?.content ?? '').join('');
        if (candidates.length !== 1 || !(patch.title ? titleMatches : same(actual, patch.elements))) throw new Error(`${file.name} ${patch.location} 回读内容不符，请人工核对`);
      } else if (file.type === 'sheet') {
        const range = `${patch.sheetId}!${patch.cell}:${patch.cell}`;
        const data = await this.api.call({ method: 'GET', url: `/open-apis/sheets/v2/spreadsheets/${enc(file.token)}/values/${range}` });
        const actual = (data.valueRange as { values?: unknown[][] } | undefined)?.values?.[0]?.[0];
        if (actual !== patch.after) throw new Error(`${file.name} ${patch.location} 回读内容不符，请人工核对`);
      } else {
        const data = await this.api.call({ method: 'GET', url: `/open-apis/bitable/v1/apps/${enc(file.token)}/tables/${enc(patch.tableId!)}/records/${enc(patch.recordId!)}` });
        const value = (data.record as { fields?: Record<string, unknown> } | undefined)?.fields?.[patch.fieldName!];
        const actual = typeof value === 'string' ? value : Array.isArray(value) ? value.map(v => typeof v?.text === 'string' ? v.text : '').join('') : '';
        if (actual !== patch.after) throw new Error(`${file.name} ${patch.location} 回读内容不符，请人工核对`);
      }
    }
  }

  private async assertSheetMerges(file: ScannedFile): Promise<void> {
    const patches = file.patches.filter(p => p.sheetId && p.cell);
    if (!patches.length) return;
    const data = await this.api.call({ method: 'GET', url: `/open-apis/sheets/v3/spreadsheets/${enc(file.token)}/sheets/query` });
    const sheets = (data.sheets as Sheet[] | undefined) ?? [];
    for (const patch of patches) {
      const sheet = sheets.find(item => item.sheet_id === patch.sheetId);
      if (!sheet) throw new Error(`${file.name} 的工作表已变化，已停止写入`);
      const cell = /^([A-Z]+)([1-9]\d*)$/.exec(patch.cell!);
      if (!cell) throw new Error(`${file.name} 的单元格坐标无效，已停止写入`);
      const merge = mergeAt(sheet, Number(cell[2]) - 1, columnIndex(cell[1]));
      if (merge?.range !== patch.mergeRange || (merge && merge.anchor !== patch.cell)) throw new Error(`${file.name} ${patch.location} 的合并区域已变化，已停止写入`);
    }
  }
}
