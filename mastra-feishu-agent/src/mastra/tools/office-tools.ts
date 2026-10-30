import { createTool } from '@mastra/core/tools';
import { z } from 'zod';
import type { Caller } from '../../identity';
import { officeAction, type OfficeAction } from '../../office/actions';
import { getOfficeApi } from '../../runtime/office-bridge';
import { requestOfficeApproval } from '../../runtime/approval-bridge';
import { requestCalendarPermission } from '../../office/permission-choice';
import { reconcileRelativeCalendarDate } from '../../office/calendar-time';
import { tableIntent, type TableIntent } from '../../office/table-service';
import { TableService } from '../../office/table-service';
import { requestTableOperation } from '../../office/table-permission';
import { ReleaseMaterialsService, parseDriveFolderUrl, releaseInput } from '../../office/release-materials';
import { sendOfficeText } from '../../runtime/office-bridge';
import { stateDb } from '../../state/store';

const token = z.string().min(1).max(256).regex(/^[\w@.-]+$/);
const time = z.string().regex(/^\d{10}$/);

function callerFrom(context: { requestContext: { get: (key: string) => unknown } }): Caller {
  const value = context.requestContext.get('caller') as Caller | undefined;
  if (!value?.userId || !value.chatId || !value.chatType) throw new Error('Trusted Feishu caller context is required');
  return value;
}

function actionTool<T extends OfficeAction['kind']>(kind: T, description: string) {
  const schema = officeAction.options.find(option => option.shape.kind.value === kind)!;
  return createTool({
    id: `office_${kind}`,
    description,
    inputSchema: (schema as z.ZodObject<z.ZodRawShape>).omit({ kind: true }),
    execute: async (input, context) => {
      const caller = callerFrom(context);
      let action = { ...input, kind } as OfficeAction;
      if (action.kind === 'doc_update' && !action.documentId.startsWith('doxcn')) {
        throw new Error('修改新版文档需要真实的 doxcn documentId；Wiki 节点或表格 token 不能用于 Docx API');
      }
      if (action.kind === 'calendar_create') {
        const originalMessage = context.requestContext.get('originalMessage');
        const messageCreatedAt = context.requestContext.get('messageCreatedAt');
        if (typeof originalMessage === 'string' && typeof messageCreatedAt === 'number') {
          action = reconcileRelativeCalendarDate(action, originalMessage, messageCreatedAt);
        }
      }
      if (action.kind === 'calendar_create' || action.kind === 'calendar_update') return requestCalendarPermission(caller, action);
      return requestOfficeApproval(caller, action);
    },
  });
}

export const readDocumentTool = createTool({
  id: 'office_read_document', description: 'Read a Feishu Docx document from its full URL or verified doxcn ID. Wiki URLs are resolved first; Sheets and Base URLs must use inspectTableTool.',
  inputSchema: z.object({ url: z.string().url().optional(), documentId: token.optional() }),
  execute: async ({ url, documentId }) => {
    const api = getOfficeApi();
    if (!url) {
      if (!documentId || !documentId.startsWith('doxcn')) throw new Error('请提供完整 Docx 链接；Wiki 节点 token 不能作为 Docx documentId');
      return api.readDocument(documentId);
    }
    const resource = await new TableService(api).resolve(url);
    if (resource.type !== 'docx') throw new Error('该链接是电子表格或多维表格，请改用表格工具');
    return api.readDocument(resource.token);
  },
});
export const readCalendarTool = createTool({
  id: 'office_read_calendar', description: 'List events in an app-accessible calendar between Unix timestamps.',
  inputSchema: z.object({ calendarId: token, startTime: time, endTime: time }),
  execute: async ({ calendarId, startTime, endTime }) => getOfficeApi().readCalendar(calendarId, startTime, endTime),
});
export const readTaskTool = createTool({
  id: 'office_read_task', description: 'Read an app-accessible Feishu task by task GUID.',
  inputSchema: z.object({ taskId: token }),
  execute: async ({ taskId }) => getOfficeApi().readTask(taskId),
});

export const createDocumentTool = actionTool('doc_create', 'Propose creating a Feishu document. Requires requester approval.');
export const updateDocumentTool = actionTool('doc_update', 'Propose replacing the text of one Feishu document block. Requires requester approval.');
export const createCalendarTool = actionTool('calendar_create', 'Create a calendar event. Omit calendarId for the requester personal calendar. Set permissionMode=user for personal calendar, app for a shared calendar. If user permission is missing, ask the requester to choose and authorize before sending any write approval card.');
export const updateCalendarTool = actionTool('calendar_update', 'Update a calendar event. Set permissionMode=user when modifying the requester personal calendar, app for an app-accessible shared calendar. Missing user permission triggers a question before approval.');
export const createTaskTool = actionTool('task_create', 'Propose creating a Feishu task. Requires requester approval.');
export const updateTaskTool = actionTool('task_update', 'Propose updating a Feishu task. Requires requester approval.');
export const sendMessageTool = actionTool('message_send', 'Propose sending a message to another allowlisted Feishu chat. Only users in FEISHU_ALLOWED_USERS may use this tool; requester approval is required.');

function tableTool<T extends TableIntent['operation']>(operation: T, description: string) {
  const schema = tableIntent.options.find(option => option.shape.operation.value === operation)!;
  return createTool({
    id: `office_table_${operation}`,
    description,
    inputSchema: (schema as z.ZodObject<z.ZodRawShape>).omit({ operation: true }),
    execute: async (input, context) => requestTableOperation(callerFrom(context), { ...input, operation } as TableIntent),
  });
}
export const inspectTableTool = tableTool('inspect', 'Resolve a Feishu Wiki, Sheets, or Base URL and inspect worksheets, fields, and a bounded preview. Use this before table edits.');
export const fillFirstEmptySheetCellTool = tableTool('fill_first_empty', 'Propose writing a value to the first empty cell beneath a named header in a Feishu spreadsheet. Requires approval.');
export const appendSheetRowTool = tableTool('append_row', 'Propose appending a complete row to a Feishu spreadsheet by header names. Requires approval.');
export const updateSheetCellTool = tableTool('update_cell', 'Propose updating a Feishu spreadsheet cell by A1 address or a unique matching row and target header. Requires approval.');
export const createBitableRecordTool = tableTool('create_record', 'Propose creating one Feishu Base record with text, number, or checkbox fields. Requires approval.');
export const updateBitableRecordTool = tableTool('update_record', 'Propose changing fields of one Feishu Base record by record ID or unique field match. Requires approval.');

export const listDriveFolderTool = createTool({
  id: 'office_drive_list_folder',
  description: 'List only the current level of a Feishu ordinary Drive folder from its /drive/folder/ URL.',
  inputSchema: z.object({ folderUrl: z.string().url() }),
  execute: async ({ folderUrl }) => new ReleaseMaterialsService(getOfficeApi()).listFolder(parseDriveFolderUrl(folderUrl)),
});
export const createDriveFolderTool = createTool({
  id: 'office_drive_create_folder',
  description: 'Propose creating one child folder in an ordinary Feishu Drive folder. Requires requester approval.',
  inputSchema: z.object({ parentFolderUrl: z.string().url(), name: z.string().min(1).max(100) }),
  execute: async ({ parentFolderUrl, name }, context) => requestOfficeApproval(callerFrom(context), { kind: 'drive_folder_create', parentFolderToken: parseDriveFolderUrl(parentFolderUrl), name }),
});
export const copyDriveFileTool = createTool({
  id: 'office_drive_copy_file',
  description: 'Propose copying one Docx, Sheet, or Base file to an ordinary Drive folder. Requires requester approval.',
  inputSchema: z.object({ fileToken: token, fileType: z.enum(['docx', 'sheet', 'bitable']), name: z.string().min(1).max(255), destinationFolderUrl: z.string().url() }),
  execute: async ({ fileToken, fileType, name, destinationFolderUrl }, context) => requestOfficeApproval(callerFrom(context), { kind: 'drive_file_copy', fileToken, fileType, name, destinationFolderToken: parseDriveFolderUrl(destinationFolderUrl) }),
});
export const renameDriveFileTool = actionTool('drive_file_rename', 'Propose renaming a Docx, Sheets, or Base file in an ordinary Drive folder. Requires requester approval and checks the current title again before writing.');
export const prepareReleaseMaterialsTool = createTool({
  id: 'office_release_materials',
  description: 'Preflight a whole release: scan current-level Docx, Sheets and Base templates, list every placeholder, then send one batch approval card. Require both ordinary Drive folder links, release ID and explicit field values. This tool does not write before approval.',
  inputSchema: releaseInput,
  execute: async (input, context) => requestReleaseMaterials(input, callerFrom(context), context.requestContext.get('messageCreatedAt'), context.requestContext.get('originalMessage')),
});

export async function requestReleaseMaterials(input: z.infer<typeof releaseInput>, caller: Caller, createdAt?: unknown, originalMessage?: unknown) {
    if (typeof originalMessage === 'string' && !originalMessage.includes(input.releaseId)) {
      throw new Error(`当前消息未包含下发编号 ${input.releaseId}；不能沿用历史会话中的批次发起审批`);
    }
    const service = new ReleaseMaterialsService(getOfficeApi());
    const manifest = await service.prepare(input, caller);
    const job = await service.createJob(caller, manifest, typeof createdAt === 'number' ? createdAt : Date.now());
    if (job.reused) return { status: job.status, approvalId: job.approvalId, jobId: job.jobId };
    const jobId = job.jobId;
    const previews = [
      `下发材料预览：${input.releaseId}\n模板文件夹：${input.templateFolderUrl}\n材料父目录：${input.materialsParentUrl}\n将创建 ${manifest.files.length} 个副本；原模板不会修改。`,
      ...manifest.files.map((file, index) => `${index + 1}. ${file.name}（${file.type}）\n${file.patches.map(p => `- ${p.location}: ${p.before} → ${p.after}`).join('\n') || '- 无占位符，原样复制'}`),
      ...(manifest.ignoredFolders.length ? [`忽略当前层子文件夹：${manifest.ignoredFolders.join('、')}`] : []),
    ];
    try {
      for (const preview of previews) for (let offset = 0; offset < preview.length; offset += 3000) await sendOfficeText(caller.chatId, preview.slice(offset, offset + 3000));
      const proposal = await requestOfficeApproval(caller, { kind: 'release_materials', jobId, releaseId: input.releaseId, fileCount: manifest.files.length, fieldSummary: Object.entries(input.fields).map(([k, v]) => `${k}=${v}`).join('；').slice(0, 1500) });
      await (await stateDb()).execute({ sql: 'UPDATE release_jobs SET approval_id = ? WHERE id = ?', args: [proposal.approvalId, jobId] });
      return proposal;
    } catch (error) {
      await (await stateDb()).execute({ sql: "UPDATE release_jobs SET status = 'failed', error = ?, updated_at = ? WHERE id = ?", args: [String(error), Date.now(), jobId] });
      throw error;
    }
}

export const officeTools = {
  readDocumentTool, readCalendarTool, readTaskTool,
  createDocumentTool, updateDocumentTool, createCalendarTool, updateCalendarTool,
  createTaskTool, updateTaskTool, sendMessageTool,
  inspectTableTool, fillFirstEmptySheetCellTool, appendSheetRowTool, updateSheetCellTool,
  createBitableRecordTool, updateBitableRecordTool,
  listDriveFolderTool, createDriveFolderTool, copyDriveFileTool, renameDriveFileTool, prepareReleaseMaterialsTool,
};
