import { withUserAccessToken, type Client, type LarkChannel } from '@larksuiteoapi/node-sdk';
import { config } from '../config';
import type { OfficeAction } from './actions';
import { TableService } from './table-service';
import { ReleaseMaterialsService } from './release-materials';

type ApiResult = { code?: number; msg?: string; data?: Record<string, unknown>; error?: { log_id?: string; field_violations?: Array<{ field?: string; description?: string }> } };
export type ApiRequest = { method: 'GET' | 'POST' | 'PATCH' | 'PUT'; url: string; data?: unknown; params?: Record<string, unknown>; userAccessToken?: string };
export type OfficeTransport = { request(request: ApiRequest): Promise<ApiResult>; send(chatId: string, text: string): Promise<unknown> };

export function createOfficeTransport(channel: LarkChannel): OfficeTransport {
  return {
    request: ({ userAccessToken, ...request }) => channel.rawClient.request(request as Parameters<Client['request']>[0], userAccessToken ? withUserAccessToken(userAccessToken) : undefined),
    send: (chatId, text) => channel.send(chatId, { text }),
  };
}

function id(value: string) { return encodeURIComponent(value); }

function safeEndpoint(url: string): string {
  return url
    .replace(/(\/spreadsheets\/)[^/]+/g, '$1:spreadsheet_token')
    .replace(/(\/values\/)[^/]+/g, '$1:range')
    .replace(/(\/sheets\/)[^/]+(?=\/(?:find|replace|filter|float_images)(?:\/|$))/g, '$1:sheet_id')
    .replace(/(\/(?:files|folders|permissions|documents|apps|tables|records)\/)[^/]+/g, '$1:id');
}

function apiError(request: ApiRequest, result: ApiResult): Error {
  const violations = result.error?.field_violations?.map(v => `${v.field ?? 'field'}: ${v.description ?? 'invalid'}`).join('; ');
  const details = [`params=${Object.keys(request.params ?? {}).join(',') || 'none'}`, violations, result.error?.log_id && `log_id=${result.error.log_id}`].filter(Boolean).join('; ');
  return new Error(`Feishu API ${result.code}: ${result.msg || 'request failed'} [${request.method} ${safeEndpoint(request.url)}${details ? `; ${details}` : ''}]`);
}

export class OfficeApi {
  constructor(private readonly transport: OfficeTransport) {}

  private async api(request: ApiRequest): Promise<Record<string, unknown>> {
    let response: ApiResult;
    try {
      response = await this.transport.request(request);
    } catch (error) {
      const response = (error as { response?: { status?: number; data?: ApiResult } })?.response;
      const apiCode = response?.data?.code;
      if (typeof apiCode === 'number' && apiCode !== 0) {
        throw apiError(request, response!.data!);
      }
      const status = response?.status
        ?? (error as { status?: number })?.status;
      if (status === 404 || status === 405 || status === 501) {
        throw new Error(`This private Feishu deployment does not provide ${request.method} ${safeEndpoint(request.url)}`);
      }
      const original = error instanceof Error ? error.message : String(error);
      throw new Error(`${original.replaceAll(request.url, safeEndpoint(request.url))} [${request.method} ${safeEndpoint(request.url)}; params=${Object.keys(request.params ?? {}).join(',') || 'none'}]`, { cause: error });
    }
    if (response.code !== undefined && response.code !== 0) {
      if ([404, 405, 501].includes(response.code)) {
        throw new Error(`This private Feishu deployment does not provide ${request.method} ${safeEndpoint(request.url)}: ${response.msg || response.code}`);
      }
      throw apiError(request, response);
    }
    return response.data ?? {};
  }

  // Table APIs use the same error handling and configured private Feishu host.
  async call(request: ApiRequest): Promise<Record<string, unknown>> { return this.api(request); }

  async readDocument(documentId: string) {
    return this.api({ method: 'GET', url: `/open-apis/docx/v1/documents/${id(documentId)}/raw_content` });
  }

  async readCalendar(calendarId: string, startTime: string, endTime: string) {
    return this.api({ method: 'GET', url: `/open-apis/calendar/v4/calendars/${id(calendarId)}/events`, params: { start_time: startTime, end_time: endTime } });
  }

  async primaryCalendar(userAccessToken: string): Promise<string> {
    const data = await this.api({ method: 'POST', url: '/open-apis/calendar/v4/calendars/primary', userAccessToken });
    const entries = data.calendars as Array<{ calendar?: { calendar_id?: string; type?: string; role?: string } }> | undefined;
    const calendar = entries?.find(entry => entry.calendar?.type === 'primary')?.calendar ?? entries?.[0]?.calendar;
    if (!calendar?.calendar_id || !['owner', 'writer'].includes(calendar.role ?? '')) throw new Error('用户主日历不存在或没有写入权限');
    return calendar.calendar_id;
  }

  async assertCalendarWritable(calendarId: string, userAccessToken: string): Promise<void> {
    const data = await this.api({ method: 'GET', url: `/open-apis/calendar/v4/calendars/${id(calendarId)}`, userAccessToken });
    const calendar = data.calendar as { role?: string } | undefined;
    if (!calendar || !['owner', 'writer'].includes(calendar.role ?? '')) throw new Error('当前用户没有该日历的写入权限');
  }

  async readTask(taskId: string) {
    return this.api({ method: 'GET', url: `/open-apis/task/v2/tasks/${id(taskId)}` });
  }

  async execute(action: OfficeAction, executionKey?: string, userAccessToken?: string): Promise<unknown> {
    switch (action.kind) {
      case 'doc_create': {
        const created = await this.api({ method: 'POST', url: '/open-apis/docx/v1/documents', data: { title: action.title, folder_token: action.folderToken } });
        if (!action.content) return created;
        const document = created.document as { document_id?: string } | undefined;
        const documentId = document?.document_id;
        if (!documentId) throw new Error('Document was created but Feishu did not return its ID; check before retrying');
        try {
          await this.api({ method: 'POST', url: `/open-apis/docx/v1/documents/${id(documentId)}/blocks/${id(documentId)}/children`, data: {
            children: [{ block_type: 2, text: { elements: [{ text_run: { content: action.content } }] } }],
          } });
        } catch (error) {
          throw new Error(`Document ${documentId} was created, but adding content failed: ${String(error)}`);
        }
        return created;
      }
      case 'doc_update':
        if (!action.documentId.startsWith('doxcn')) throw new Error('Wiki 节点或表格 token 不能作为 Docx documentId');
        return this.api({ method: 'PATCH', url: `/open-apis/docx/v1/documents/${id(action.documentId)}/blocks/${id(action.blockId)}`, data: {
          update_text_elements: { elements: [{ text_run: { content: action.text } }] },
        } });
      case 'calendar_create':
        if (Number(action.startTime) >= Number(action.endTime)) throw new Error('Calendar end time must be after start time');
        if (!action.calendarId) throw new Error('创建日程前必须确定日历 ID');
        if (action.permissionMode === 'user' && !userAccessToken) throw new Error('缺少请求者的用户日历权限');
        return this.api({ method: 'POST', url: `/open-apis/calendar/v4/calendars/${id(action.calendarId)}/events`, data: {
          summary: action.summary, description: action.description,
          start_time: { timestamp: action.startTime }, end_time: { timestamp: action.endTime },
        }, userAccessToken: action.permissionMode === 'user' ? userAccessToken : undefined });
      case 'calendar_update': {
        if (action.permissionMode === 'user' && !userAccessToken) throw new Error('缺少请求者的用户日历权限');
        const data: Record<string, unknown> = {};
        if (action.summary !== undefined) data.summary = action.summary;
        if (action.description !== undefined) data.description = action.description;
        if (action.startTime !== undefined) data.start_time = { timestamp: action.startTime };
        if (action.endTime !== undefined) data.end_time = { timestamp: action.endTime };
        if (!Object.keys(data).length) throw new Error('No calendar fields to update');
        return this.api({ method: 'PATCH', url: `/open-apis/calendar/v4/calendars/${id(action.calendarId)}/events/${id(action.eventId)}`, data, userAccessToken: action.permissionMode === 'user' ? userAccessToken : undefined });
      }
      case 'task_create':
        return this.api({ method: 'POST', url: '/open-apis/task/v2/tasks', data: {
          summary: action.summary, description: action.description,
          due: action.dueTimestamp ? { timestamp: action.dueTimestamp } : undefined,
          client_token: executionKey,
        } });
      case 'task_update': {
        const task: Record<string, unknown> = {};
        if (action.summary !== undefined) task.summary = action.summary;
        if (action.description !== undefined) task.description = action.description;
        if (action.dueTimestamp !== undefined) task.due = { timestamp: action.dueTimestamp };
        if (!Object.keys(task).length) throw new Error('No task fields to update');
        return this.api({ method: 'PATCH', url: `/open-apis/task/v2/tasks/${id(action.taskId)}`, data: { task, update_fields: Object.keys(task) } });
      }
      case 'message_send':
        if (!config.allowedChats.has(action.targetChatId)) throw new Error('Target chat is not allowlisted');
        return this.transport.send(action.targetChatId, action.text);
      case 'sheet_cell_set':
      case 'sheet_row_append':
      case 'bitable_record_create':
      case 'bitable_record_update':
        return new TableService(this).execute(action, userAccessToken);
      case 'drive_folder_create':
        return new ReleaseMaterialsService(this).createFolder(action.parentFolderToken, action.name);
      case 'drive_file_copy':
        return new ReleaseMaterialsService(this).copyFile({ token: action.fileToken, name: action.name, type: action.fileType }, action.destinationFolderToken);
      case 'drive_file_rename':
        return new ReleaseMaterialsService(this).renameFile(action.fileToken, action.fileType, action.folderToken, action.expectedName, action.newName);
      case 'release_materials':
        throw new Error('整批材料须通过 release-materials Workflow 执行');
    }
  }
}
