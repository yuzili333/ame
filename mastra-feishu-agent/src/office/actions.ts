import { z } from 'zod';
import { formatShanghaiTime } from './calendar-time';

const token = z.string().min(1).max(256).regex(/^[\w@.-]+$/);
const text = z.string().trim().min(1).max(20_000);
const calendarTimestamp = z.string().regex(/^\d{10}$/, 'Calendar timestamps must be Unix seconds');
const taskTimestamp = z.string().regex(/^\d{13}$/, 'Task timestamps must be Unix milliseconds');
const cellValue = z.union([z.string().max(40_000), z.number(), z.boolean()]);
const fields = z.record(z.string().min(1).max(255), cellValue).refine(value => Object.keys(value).length > 0, '至少提供一个字段');
const tableIdentity = z.enum(['app', 'user']);

export const officeAction = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('doc_create'), title: text.max(255), content: text.optional(), folderToken: token.optional() }),
  z.object({ kind: z.literal('doc_update'), documentId: token, blockId: token, text }),
  z.object({ kind: z.literal('calendar_create'), calendarId: token.optional(), permissionMode: z.enum(['auto', 'user', 'app']).optional(), summary: text.max(255), description: text.optional(), startTime: calendarTimestamp, endTime: calendarTimestamp }),
  z.object({ kind: z.literal('calendar_update'), calendarId: token, permissionMode: z.enum(['user', 'app']).optional(), eventId: token, summary: text.max(255).optional(), description: text.optional(), startTime: calendarTimestamp.optional(), endTime: calendarTimestamp.optional() }),
  z.object({ kind: z.literal('task_create'), summary: text.max(255), description: text.optional(), dueTimestamp: taskTimestamp.optional() }),
  z.object({ kind: z.literal('task_update'), taskId: token, summary: text.max(255).optional(), description: text.optional(), dueTimestamp: taskTimestamp.optional() }),
  z.object({ kind: z.literal('message_send'), targetChatId: token, text }),
  z.object({ kind: z.literal('sheet_cell_set'), spreadsheetToken: token, sheetId: token, cell: z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/), expectedValue: cellValue.nullable(), value: cellValue, permissionMode: tableIdentity }),
  z.object({ kind: z.literal('sheet_row_append'), spreadsheetToken: token, sheetId: token, headers: z.array(z.string()).min(1).max(256), values: z.array(cellValue.nullable()).min(1).max(256), permissionMode: tableIdentity }),
  z.object({ kind: z.literal('bitable_record_create'), appToken: token, tableId: token, fields, permissionMode: tableIdentity }),
  z.object({ kind: z.literal('bitable_record_update'), appToken: token, tableId: token, recordId: token, fields, expectedFields: z.record(z.string(), z.unknown()), permissionMode: tableIdentity }),
  z.object({ kind: z.literal('drive_folder_create'), parentFolderToken: token, name: text.max(100) }),
  z.object({ kind: z.literal('drive_file_copy'), fileToken: token, fileType: z.enum(['docx', 'sheet', 'bitable']), name: text.max(255), destinationFolderToken: token }),
  z.object({ kind: z.literal('drive_file_rename'), fileToken: token, fileType: z.enum(['docx', 'sheet', 'bitable']), folderToken: token, expectedName: text.max(255), newName: text.max(255) }),
  z.object({ kind: z.literal('release_materials'), jobId: z.string().uuid(), releaseId: text.max(100), fileCount: z.number().int().positive(), fieldSummary: text }),
]);

export type OfficeAction = z.infer<typeof officeAction>;

export function actionSummary(action: OfficeAction): string {
  switch (action.kind) {
    case 'doc_create': return `创建文档：${action.title}${action.content ? `；正文：${action.content.slice(0, 180)}` : ''}`;
    case 'doc_update': return `修改文档 ${action.documentId} 的块 ${action.blockId}：${action.text.slice(0, 180)}`;
    case 'calendar_create': return `以${action.permissionMode === 'user' ? '用户' : '应用'}身份在日历 ${action.calendarId ?? '未指定'} 创建日程：${action.summary}（北京时间 ${formatShanghaiTime(action.startTime)}—${formatShanghaiTime(action.endTime)}）`;
    case 'calendar_update': return `以${action.permissionMode === 'user' ? '用户' : '应用'}身份修改日程 ${action.eventId}：${action.summary ?? '时间或说明'}${action.startTime || action.endTime ? `（北京时间 ${action.startTime ? formatShanghaiTime(action.startTime) : '原开始时间'}—${action.endTime ? formatShanghaiTime(action.endTime) : '原结束时间'}）` : ''}`;
    case 'task_create': return `创建任务：${action.summary}`;
    case 'task_update': return `修改任务 ${action.taskId}：${action.summary ?? '说明或截止时间'}`;
    case 'message_send': return `向会话 ${action.targetChatId} 发送：${action.text.slice(0, 180)}`;
    case 'sheet_cell_set': return `以${action.permissionMode === 'user' ? '用户' : '应用'}身份修改电子表格 ${action.spreadsheetToken} / ${action.sheetId} 的 ${action.cell}：${JSON.stringify(action.expectedValue)} → ${JSON.stringify(action.value)}`;
    case 'sheet_row_append': return `以${action.permissionMode === 'user' ? '用户' : '应用'}身份在电子表格 ${action.spreadsheetToken} / ${action.sheetId} 追加一行：${JSON.stringify(Object.fromEntries(action.headers.map((header, index) => [header, action.values[index]]))).slice(0, 1200)}`;
    case 'bitable_record_create': return `以${action.permissionMode === 'user' ? '用户' : '应用'}身份在多维表格 ${action.appToken} / ${action.tableId} 新增记录：${JSON.stringify(action.fields).slice(0, 1200)}`;
    case 'bitable_record_update': return `以${action.permissionMode === 'user' ? '用户' : '应用'}身份修改多维表格 ${action.appToken} / ${action.tableId} 的记录 ${action.recordId}：${JSON.stringify(action.expectedFields).slice(0, 600)} → ${JSON.stringify(action.fields).slice(0, 600)}`;
    case 'drive_folder_create': return `在文件夹 ${action.parentFolderToken} 创建子文件夹“${action.name}”`;
    case 'drive_file_copy': return `将 ${action.fileType} 文件“${action.name}”复制到文件夹 ${action.destinationFolderToken}`;
    case 'drive_file_rename': return `在文件夹 ${action.folderToken} 中将 ${action.fileType} 文件 ${action.fileToken} 的标题由“${action.expectedName}”改为“${action.newName}”`;
    case 'release_materials': return `整批生成下发材料 ${action.releaseId}：${action.fileCount} 个文件；字段 ${action.fieldSummary}`;
  }
}

export function actionResultDetail(action: OfficeAction, result: unknown): string | undefined {
  if (action.kind === 'drive_file_rename') return `${actionSummary(action)}\n标题已回读确认。`;
  if (action.kind === 'release_materials') {
    const release = result as { folderUrl?: string; files?: Array<{ name: string; url: string }> } | null;
    return `材料文件夹：${release?.folderUrl ?? '未返回'}\n${release?.files?.map(file => `- ${file.name}: ${file.url}`).join('\n') ?? ''}`;
  }
  if (action.kind === 'sheet_cell_set' || action.kind === 'sheet_row_append' || action.kind === 'bitable_record_create' || action.kind === 'bitable_record_update') {
    const recordId = (result as { record?: { record_id?: unknown } } | null)?.record?.record_id;
    return `${actionSummary(action)}${typeof recordId === 'string' ? `\n\n记录 ID：${recordId}` : ''}`;
  }
  if (action.kind !== 'calendar_create') return undefined;
  const event = (result as { event?: { event_id?: unknown; app_link?: unknown } } | null)?.event;
  const lines = [actionSummary(action)];
  if (typeof event?.event_id === 'string') lines.push(`事件 ID：${event.event_id}`);
  if (typeof event?.app_link === 'string') {
    try {
      const link = new URL(event.app_link);
      if (link.protocol === 'https:') lines.push(`[打开日程](${link.toString()})`);
    } catch { /* Ignore an invalid link returned by the remote API. */ }
  }
  return lines.join('\n\n');
}
