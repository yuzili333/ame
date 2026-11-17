import { createTool } from '@mastra/core/tools';
import { z } from 'zod';
import { config } from '../../config';
import type { Caller } from '../../identity';
import { officeAction, type OfficeAction } from '../../office/actions';
import { getOfficeApi } from '../../runtime/office-bridge';
import { requestOfficeApproval } from '../../runtime/approval-bridge';
import { requestResourceOperation } from '../../office/resource-permission';
import { isResourceWrite, resourceRead, type ResourceRead } from '../../office/resource-operation';
import { reconcileRelativeCalendarDate } from '../../office/calendar-time';
import { tableIntent, type TableIntent } from '../../office/table-service';
import { TableService } from '../../office/table-service';
import { requestTableOperation } from '../../office/table-permission';
import { ReleaseMaterialsService, parseDriveFolderUrl, releaseInput } from '../../office/release-materials';
import { sendOfficeText } from '../../runtime/office-bridge';
import { isMessageRecalled, stateDb } from '../../state/store';

const token = z.string().min(1).max(256).regex(/^[\w@.-]+$/);

function callerFrom(context: { requestContext: { get: (key: string) => unknown } }): Caller {
  const value = context.requestContext.get('caller') as Caller | undefined;
  if (!value?.userId || !value.chatId || !value.chatType) throw new Error('Trusted Feishu caller context is required');
  return value;
}
function continuationFrom(context: { requestContext: { get: (key: string) => unknown } }) {
  const originalMessage = context.requestContext.get('originalMessage');
  const messageCreatedAt = context.requestContext.get('messageCreatedAt');
  return typeof originalMessage === 'string' && typeof messageCreatedAt === 'number' ? { originalMessage, messageCreatedAt } : undefined;
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
      if (isResourceWrite(action)) return requestResourceOperation(caller, action, continuationFrom(context));
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
function resourceReadTool(kind: ResourceRead['kind'], id: string, description: string) {
  const schema = resourceRead.options.find(option => option.shape.kind.value === kind)!;
  return createTool({ id, description, inputSchema: (schema as z.ZodObject<z.ZodRawShape>).omit({ kind: true }),
    execute: async (input, context) => requestResourceOperation(callerFrom(context), { ...input, kind } as ResourceRead, continuationFrom(context)) });
}
export const readCalendarTool = resourceReadTool('calendar_events', 'office_read_calendar', 'List events in a personal or shared calendar in Unix seconds. Omit calendarId for the requester primary calendar. Default user identity; explicitly use app for robot-accessible shared calendars. Supports pageToken. Missing authorization sends a consent card and resumes this query after OAuth.');
export const listCalendarsTool = resourceReadTool('calendar_list', 'office_list_calendars', 'List personal and subscribed shared calendars visible to the requester, including IDs, types and roles. Supports pagination. Use app only for robot calendars.');
export const readCalendarInfoTool = resourceReadTool('calendar_read', 'office_read_calendar_info', 'Read calendar entity metadata and access role. Omit calendarId for the requester primary calendar.');
export const readCalendarEventTool = resourceReadTool('calendar_event_read', 'office_read_calendar_event', 'Read one event by verified calendarId and eventId using requester identity, or explicit app identity.');
export const readTaskTool = resourceReadTool('task_read', 'office_read_task', 'Read the requester task by verified task GUID. Missing user authorization sends a consent card and resumes after OAuth.');
export const listTasksTool = resourceReadTool('task_list', 'office_list_tasks', 'List the requester assigned tasks (我负责的), optionally filter completed and continue with pageToken. This API does not list every task in the tenant.');

export const createDocumentTool = actionTool('doc_create', 'Propose creating a Feishu document. Requires requester approval.');
export const updateDocumentTool = actionTool('doc_update', 'Propose replacing the text of one Feishu document block. Requires requester approval.');
export const createCalendarTool = actionTool('calendar_create', 'Create an event, not a calendar entity. Omit calendarId for the requester primary calendar. Default user identity, app only for robot-accessible shared calendars. Missing permission triggers authorization consent before operation approval.');
export const updateCalendarTool = actionTool('calendar_update', 'Update an event in the selected personal or shared calendar. Default user identity. Missing permission triggers authorization consent; writing requires requester approval.');
export const deleteCalendarTool = actionTool('calendar_delete', 'Delete one event by verified calendarId and eventId. Requires requester approval; missing permissions trigger authorization consent.');
export const createCalendarContainerTool = actionTool('calendar_container_create', 'Create a shared calendar entity under the requester identity, or explicit app identity. The personal primary calendar cannot be created via this API. Requires approval.');
export const updateCalendarContainerTool = actionTool('calendar_container_update', 'Edit the selected calendar entity title, description or visibility; requires owner role and requester approval.');
export const deleteCalendarContainerTool = actionTool('calendar_container_delete', 'Delete the entire explicitly selected shared calendar including its events. Requires owner role and requester approval. Never use for deleting an event or the personal primary calendar.');
export const createTaskTool = actionTool('task_create', 'Create a task as the requester (default user identity). startTimestamp and dueTimestamp are Unix milliseconds; a requested time range must populate both, not only the description. Missing authorization sends a consent card, OAuth resumes the request, then operation approval is required.');
export const updateTaskTool = actionTool('task_update', 'Edit a requester task. startTimestamp and dueTimestamp are Unix milliseconds for the start and deadline. completedAt is Unix milliseconds, or 0 to reopen. Missing authorization triggers consent; writing requires approval.');
export const deleteTaskTool = actionTool('task_delete', 'Delete the explicitly selected requester task by task GUID. Missing authorization triggers consent; deletion requires approval.');
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
const loggedReleaseErrors = new WeakSet<object>();
function logReleaseError(input: z.infer<typeof releaseInput>, caller: Caller, phase: string, error: unknown) {
  if (error && typeof error === 'object') {
    if (loggedReleaseErrors.has(error)) return;
    loggedReleaseErrors.add(error);
  }
  let detail = error instanceof Error ? error.message : String(error);
  for (const secret of [config.appSecret, config.modelApiKey, config.tokenEncryptionKey, input.templateFolderUrl, input.materialsParentUrl]) {
    if (secret) detail = detail.replaceAll(secret, '[REDACTED]');
  }
  console.error('[prepareReleaseMaterialsTool] failed', JSON.stringify({ phase, releaseId: input.releaseId, chatType: caller.chatType, error: detail.slice(0, 1200) }));
}
export const prepareReleaseMaterialsTool = createTool({
  id: 'office_release_materials',
  description: 'Preflight a whole release from current-level Docx, Sheets and Base templates, then send one batch approval card. Require both ordinary Drive folder links and a 12-digit release ID. Supply fields as a dynamic map of only the field values explicitly requested in this message; release ID automatically fills {{下发编号}}. Unspecified placeholders remain unchanged, and sheets may have different field subsets. For a large Sheets grid, ask for optional sheetScopes entries with exact fileName, sheetName and A1:B20 range; only those worksheets and ranges will be scanned. This tool does not write before approval.',
  inputSchema: releaseInput,
  execute: async (input, context) => {
    const caller = callerFrom(context);
    try { return await requestReleaseMaterials(input, caller, context.requestContext.get('messageCreatedAt'), context.requestContext.get('originalMessage')); }
    catch (error) { logReleaseError(input, caller, 'tool', error); throw error; }
  },
});

export async function requestReleaseMaterials(input: z.infer<typeof releaseInput>, caller: Caller, createdAt?: unknown, originalMessage?: unknown) {
    if (typeof originalMessage === 'string' && !originalMessage.includes(input.releaseId)) {
      throw new Error(`当前消息未包含下发编号 ${input.releaseId}；不能沿用历史会话中的批次发起审批`);
    }
    const service = new ReleaseMaterialsService(getOfficeApi());
    let manifest: Awaited<ReturnType<ReleaseMaterialsService['prepare']>>;
    try { manifest = await service.prepare(input, caller); }
    catch (error) {
      logReleaseError(input, caller, 'preflight', error);
      throw error;
    }
    if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，本轮下发已取消');
    const matchedFields = [...new Set(manifest.files.flatMap(file => file.patches.flatMap(patch => patch.fields)))];
    const preview = [
      `下发材料预览：${input.releaseId}`,
      `将创建 ${manifest.files.length} 个副本；原模板不会修改。`,
      `实际替换字段：${matchedFields.join('、') || '无'}`,
      ...(manifest.input.sheetScopes?.length ? [`限定扫描范围：${manifest.input.sheetScopes.map(scope => `${scope.fileName}/${scope.sheetName}!${scope.range}`).join('；')}`] : []),
      ...manifest.files.map((file, index) => `${index + 1}. ${file.name}（${file.type}）：${file.patches.length ? `${file.patches.reduce((total, patch) => total + patch.count, 0)} 处；${[...new Set(file.patches.flatMap(p => p.fields))].join('、')}` : '原样复制'}`),
      ...(manifest.unusedFields?.length ? [`未使用字段：${manifest.unusedFields.join('、')}`] : []),
      ...(manifest.ignoredFolders.length ? [`忽略子文件夹：${manifest.ignoredFolders.join('、')}`] : []),
    ].join('\n');
    if (preview.length > 2800) throw new Error('预览文件清单过长，请减少模板文件数量后重新发起；尚未发送审批卡片');
    const fieldSummary = Object.entries(manifest.input.fields)
      .map(([key, value]) => `${key}=${value.length > 120 ? `${value.slice(0, 120)}…（完整值见原请求）` : value}`)
      .join('；');
    if (fieldSummary.length > 1500) throw new Error('审批字段摘要过长，请减少本批次字段数量；尚未发送审批卡片');
    const job = await service.createJob(caller, manifest, typeof createdAt === 'number' ? createdAt : Date.now());
    if (job.reused) return { status: job.status, approvalId: job.approvalId, jobId: job.jobId };
    const jobId = job.jobId;
    try {
      if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，本轮下发已取消');
      await sendOfficeText(caller.chatId, preview);
      const proposal = await requestOfficeApproval(caller, { kind: 'release_materials', jobId, releaseId: input.releaseId, fileCount: manifest.files.length, fieldSummary });
      await (await stateDb()).execute({ sql: 'UPDATE release_jobs SET approval_id = ? WHERE id = ?', args: [proposal.approvalId, jobId] });
      return proposal;
    } catch (error) {
      await (await stateDb()).execute({ sql: "UPDATE release_jobs SET status = 'failed', error = ?, updated_at = ? WHERE id = ? AND status = 'pending'", args: [String(error), Date.now(), jobId] });
      throw error;
    }
}

export const officeTools = {
  readDocumentTool, readCalendarTool, readTaskTool,
  listCalendarsTool, readCalendarInfoTool, readCalendarEventTool, listTasksTool,
  createCalendarContainerTool, updateCalendarContainerTool, deleteCalendarContainerTool, deleteCalendarTool, deleteTaskTool,
  createDocumentTool, updateDocumentTool, createCalendarTool, updateCalendarTool,
  createTaskTool, updateTaskTool, sendMessageTool,
  inspectTableTool, fillFirstEmptySheetCellTool, appendSheetRowTool, updateSheetCellTool,
  createBitableRecordTool, updateBitableRecordTool,
  listDriveFolderTool, createDriveFolderTool, copyDriveFileTool, renameDriveFileTool, prepareReleaseMaterialsTool,
};
