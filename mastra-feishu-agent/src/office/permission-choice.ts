import { randomUUID } from 'node:crypto';
import type { Caller } from '../identity';
import { type OfficeAction } from './actions';
import { getOfficeApi } from '../runtime/office-bridge';
import { requestOfficeApproval } from '../runtime/approval-bridge';
import { stateDb } from '../state/store';
import { authorizationUrl, exchangeCodeForUser, invalidateUserAccessToken, userAccessToken } from './user-auth';

type CalendarAction = Extract<OfficeAction, { kind: 'calendar_create' | 'calendar_update' }>;
type Pending = { id: string; requester_id: string; chat_id: string; chat_type?: Caller['chatType']; payload: string; status: string; expires_at: number };
type Notifier = (chatId: string, text: string) => Promise<void>;
let notifier: Notifier | undefined;
export function setPermissionNotifier(value: Notifier | undefined) { notifier = value; }
async function tell(chatId: string, text: string) { if (!notifier) throw new Error('权限追问需要运行飞书网关'); await notifier(chatId, text); }

async function pendingFor(caller: Caller): Promise<Pending | null> {
  const rows = (await (await stateDb()).execute({ sql: `SELECT * FROM permission_requests WHERE requester_id = ? AND chat_id = ?
    AND status IN ('choice', 'choice_reauth', 'calendar_id', 'calendar_id_reauth', 'auth') AND expires_at > ? ORDER BY created_at DESC LIMIT 1`, args: [caller.userId, caller.chatId, Date.now()] })).rows;
  return rows[0] as unknown as Pending | undefined ?? null;
}

export async function latestPermissionRequestId(caller: Caller): Promise<string | null> {
  const rows = (await (await stateDb()).execute({ sql: `SELECT id FROM permission_requests WHERE requester_id = ? AND chat_id = ?
    ORDER BY created_at DESC, rowid DESC LIMIT 1`, args: [caller.userId, caller.chatId] })).rows;
  return rows[0]?.id as string | undefined ?? null;
}

function isCalendarAuthorizationError(error: unknown): boolean {
  return /Feishu API (403|99991672|99991676|99991677|99991679|191002|193002)|用户日历授权已过期|没有该日历的写入权限|没有写入权限/.test(String(error));
}

async function createPending(caller: Caller, action: CalendarAction, reauthorize = false) {
  const db = await stateDb();
  await db.execute({ sql: `UPDATE permission_requests SET status = 'superseded' WHERE requester_id = ? AND chat_id = ? AND status IN ('choice','choice_reauth','calendar_id','calendar_id_reauth','auth','failed','table_choice','table_auth','table_failed')`, args: [caller.userId, caller.chatId] });
  const id = randomUUID();
  await db.execute({ sql: `INSERT INTO permission_requests(id, requester_id, chat_id, chat_type, payload, status, created_at, expires_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)`, args: [id, caller.userId, caller.chatId, caller.chatType, JSON.stringify(action), reauthorize ? 'choice_reauth' : 'choice', Date.now(), Date.now() + 30 * 60_000] });
  return id;
}

async function approveResolved(caller: Caller, action: CalendarAction) {
  if (action.permissionMode === 'user') {
    const token = await userAccessToken(caller.userId);
    if (!token) throw new Error('用户日历授权尚未完成');
    if (action.kind === 'calendar_create' && !action.calendarId) action = { ...action, calendarId: await getOfficeApi().primaryCalendar(token) };
    else await getOfficeApi().assertCalendarWritable(action.calendarId!, token);
  }
  if (action.kind === 'calendar_create' && !action.calendarId) throw new Error('应用身份创建日程需要日历 ID');
  return requestOfficeApproval(caller, action);
}

export async function requestCalendarPermission(caller: Caller, action: CalendarAction) {
  const mode = action.permissionMode ?? (action.kind === 'calendar_create' && !action.calendarId ? 'auto' : 'app');
  if (mode === 'app') {
    if (action.kind !== 'calendar_create' || action.calendarId) return approveResolved(caller, { ...action, permissionMode: 'app' });
  }
  let reauthorize = false;
  if (mode === 'user') {
    let token: string | null = null;
    try {
      token = await userAccessToken(caller.userId);
      if (token) return await approveResolved(caller, { ...action, permissionMode: 'user' });
    } catch (error) {
      if (!isCalendarAuthorizationError(error)) throw error;
      if (token) await invalidateUserAccessToken(caller.userId, token);
      reauthorize = true;
    }
  }
  const requestId = await createPending(caller, action, reauthorize);
  const question = reauthorize
    ? '当前用户令牌缺少查询或写入日历所需的权限。要重新授权“用户权限”，还是改用机器人有权访问的“应用权限”？请回复“用户权限”或“应用权限”；此时尚未发起写入审批。'
    : mode === 'user'
    ? '你的个人日历尚未授权。要使用「用户权限」在你的主日历创建/修改日程，还是改用「应用权限」操作已共享给机器人的日历？请回复“用户权限”或“应用权限”；回复“取消”可放弃。此时尚未发起写入审批。'
    : '这条日程要写入你的个人主日历（用户权限），还是机器人有权访问的共享日历（应用权限）？请回复“用户权限”或“应用权限”；回复“取消”可放弃。此时尚未发起写入审批。';
  try { await tell(caller.chatId, question); }
  catch (error) {
    await (await stateDb()).execute({ sql: "UPDATE permission_requests SET status = 'failed' WHERE id = ?", args: [requestId] });
    throw error;
  }
  return { status: 'permission_choice_required', requestId, question };
}

async function sendAuthorizationLink(caller: Caller, requestId: string, action: CalendarAction): Promise<void> {
  let url: string;
  try { url = authorizationUrl(requestId); }
  catch (error) {
    await tell(caller.chatId, `${error instanceof Error ? error.message : String(error)}。当前请求没有发起审批；可回复“应用权限”改用已共享的日历，或回复“取消”。`);
    return;
  }
  await (await stateDb()).execute({ sql: "UPDATE permission_requests SET status = 'auth', payload = ? WHERE id = ?", args: [JSON.stringify(action), requestId] });
  await tell(caller.chatId, `请以你自己的飞书账号打开并同意日历授权：${url}\n授权完成后，我会先发操作审批卡片；授权本身不会创建日程。`);
}

export async function resolvePermissionReply(caller: Caller, content: string): Promise<boolean> {
  const answer = content.trim().replace(/^@\S+\s*/, '');
  const db = await stateDb();
  const pending = await pendingFor(caller);
  if (!pending) {
    if (!/^(重新授权|重新申请授权|reauthorize)$/i.test(answer)) return false;
    const failed = (await db.execute({ sql: `SELECT * FROM permission_requests WHERE requester_id = ? AND chat_id = ?
      AND status = 'failed' AND expires_at > ? ORDER BY created_at DESC, rowid DESC LIMIT 1`, args: [caller.userId, caller.chatId, Date.now()] })).rows[0] as unknown as Pending | undefined;
    if (!failed) return false;
    const action = { ...JSON.parse(failed.payload) as CalendarAction, permissionMode: 'user' as const };
    const requestId = await createPending(caller, action, true);
    await sendAuthorizationLink(caller, requestId, action);
    return true;
  }
  if (/^(取消|算了|cancel)$/i.test(answer)) {
    await db.execute({ sql: "UPDATE permission_requests SET status = 'cancelled' WHERE id = ?", args: [pending.id] });
    await tell(caller.chatId, '已取消这次日程请求。');
    return true;
  }
  const action = JSON.parse(pending.payload) as CalendarAction;
  if (pending.status.startsWith('calendar_id') && !/^(用户权限|用户身份|个人日历|user|应用权限|应用身份|app)$/i.test(answer)) {
    const calendarId = answer.match(/^[\w@.-]{1,256}$/)?.[0];
    if (!calendarId || action.kind !== 'calendar_create') return false;
    await approveResolved(caller, { ...action, calendarId, permissionMode: 'app' });
    await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [pending.id] });
    return true;
  }
  if (/^(应用权限|应用身份|app)$/i.test(answer)) {
    if (action.kind === 'calendar_create' && !action.calendarId) {
      await db.execute({ sql: 'UPDATE permission_requests SET status = ? WHERE id = ?', args: [pending.status.endsWith('_reauth') ? 'calendar_id_reauth' : 'calendar_id', pending.id] });
      await tell(caller.chatId, '请发送要写入的共享日历 ID。机器人必须对该日历有写入权限；收到 ID 后才会发审批卡片。');
    } else {
      await approveResolved(caller, { ...action, permissionMode: 'app' });
      await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [pending.id] });
    }
    return true;
  }
  if (/^(用户权限|用户身份|个人日历|user)$/i.test(answer)) {
    const selected = { ...action, permissionMode: 'user' as const };
    if (pending.status !== 'auth' && !pending.status.endsWith('_reauth')) {
      const token = await userAccessToken(caller.userId);
      if (token) {
        try {
          await approveResolved(caller, selected);
          await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [pending.id] });
          return true;
        } catch (error) {
          if (!isCalendarAuthorizationError(error)) throw error;
          await invalidateUserAccessToken(caller.userId, token);
        }
      }
    }
    await sendAuthorizationLink(caller, pending.id, selected);
    return true;
  }
  return false;
}

export async function completeCalendarAuthorization(state: string, code: string) {
  const db = await stateDb();
  const claimed = await db.execute({ sql: `UPDATE permission_requests SET status = 'auth_processing'
    WHERE id = ? AND status = 'auth' AND expires_at > ? RETURNING *`, args: [state, Date.now()] });
  const pending = claimed.rows[0] as unknown as Pending | undefined;
  if (!pending) throw new Error('授权请求已失效或已处理');
  const action = JSON.parse(pending.payload) as CalendarAction;
  const caller: Caller = { userId: pending.requester_id, chatId: pending.chat_id, chatType: pending.chat_type ?? 'p2p' };
  let exchangedToken: string | undefined;
  try {
    exchangedToken = await exchangeCodeForUser(code, caller.userId);
    await approveResolved(caller, { ...action, permissionMode: 'user' });
    await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [state] });
  } catch (error) {
    if (exchangedToken && isCalendarAuthorizationError(error)) await invalidateUserAccessToken(caller.userId, exchangedToken);
    await db.execute({ sql: "UPDATE permission_requests SET status = 'failed' WHERE id = ?", args: [state] });
    await tell(caller.chatId, `用户日历授权或审批准备失败：${error instanceof Error ? error.message : String(error)}。30 分钟内可回复“重新授权”再次申请，或重新发起日程请求。`);
    throw error;
  }
}
