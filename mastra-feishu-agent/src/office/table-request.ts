import type { TableIntent } from './table-service';

/** Route only a fully specified, single-cell request without relying on conversation memory. */
export function parseFirstEmptySheetCellRequest(content: string): Extract<TableIntent, { operation: 'fill_first_empty' }> | undefined {
  const instruction = /表头列\s*[“"「]([^”"」]{1,100})[”"」][\s\S]{0,80}?第一个空白单元格\s*写入\s*[“"「]([^”"」]{1,40000})[”"」]/.exec(content);
  if (!instruction) return undefined;
  const candidates = content.match(/https?:\/\/[^\s)）]+/gi) ?? [];
  const url = candidates.map(candidate => candidate.replace(/[\]，。；;]+$/, '')).find(candidate => {
    try { return ['wiki', 'sheets'].includes(new URL(candidate).pathname.split('/')[1]); }
    catch { return false; }
  });
  if (!url) return undefined;
  return {
    operation: 'fill_first_empty',
    url,
    columnHeader: instruction[1],
    value: instruction[2],
    permissionMode: /(?:使用|采用|以|用)\s*用户权限|用户身份/.test(content) ? 'user' : 'app',
  };
}
