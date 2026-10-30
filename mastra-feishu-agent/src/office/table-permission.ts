import { randomUUID } from 'node:crypto';
import { config } from '../config';
import type { Caller } from '../identity';
import { requestOfficeApproval } from '../runtime/approval-bridge';
import { getOfficeApi } from '../runtime/office-bridge';
import { stateDb } from '../state/store';
import { CALENDAR_SCOPES, authorizationUrl, exchangeCodeForUser, userAccessToken } from './user-auth';
import { TableService, type TableIntent } from './table-service';

type Pending = { id: string; requester_id: string; chat_id: string; chat_type?: Caller['chatType']; payload: string; status: string; expires_at: number };
type Notify = (chatId: string, message: string) => Promise<void>;
let notify: Notify | undefined;
export function setTablePermissionNotifier(value: Notify | undefined) { notify = value; }
async function tell(chatId: string, message: string) { if (!notify) throw new Error('表格权限追问需要运行飞书网关'); await notify(chatId, message); }
function service() { return new TableService(getOfficeApi()); }
const authError = (error: unknown) => /\b403\b|Feishu API (91403|99991672|99991676|99991677|99991679|1254302)|permission denied|Access denied|没有编辑权限|用户表格授权已失效/i.test(String(error));

async function requestApprovalOnce(caller: Caller, action: Awaited<ReturnType<TableService['prepare']>>) {
  const db = await stateDb();
  const payload = JSON.stringify(action);
  const existing = (await db.execute({ sql: `SELECT id, expires_at FROM actions WHERE requester_id = ? AND chat_id = ? AND kind = ? AND payload = ? AND status = 'pending' AND expires_at > ? ORDER BY created_at DESC LIMIT 1`, args: [caller.userId, caller.chatId, action.kind, payload, Date.now()] })).rows[0];
  if (existing) return { status: 'pending', approvalId: String(existing.id), expiresAt: Number(existing.expires_at) };
  return requestOfficeApproval(caller, action);
}

async function perform(caller: Caller, intent: TableIntent, mode: 'app' | 'user', token?: string) {
  if (intent.operation === 'inspect') {
    const result = await service().inspect(intent, token);
    return { status: 'read', result };
  }
  const action = await service().prepare(intent, mode, token);
  return requestApprovalOnce(caller, action);
}

async function createPending(caller: Caller, intent: TableIntent, reason: string) {
  const db = await stateDb();
  await db.execute({ sql: `UPDATE permission_requests SET status = 'superseded' WHERE requester_id = ? AND chat_id = ? AND status IN ('choice','choice_reauth','calendar_id','calendar_id_reauth','auth','table_choice','table_auth','table_failed','failed')`, args: [caller.userId, caller.chatId] });
  const id = randomUUID();
  await db.execute({ sql: `INSERT INTO permission_requests(id, requester_id, chat_id, chat_type, payload, status, created_at, expires_at) VALUES (?, ?, ?, ?, ?, 'table_choice', ?, ?)`, args: [id, caller.userId, caller.chatId, caller.chatType, JSON.stringify(intent), Date.now(), Date.now() + 30 * 60_000] });
  const question = `${reason}。是否改用你自己的飞书用户权限访问这份表格？请回复“用户权限”以授权，或回复“应用权限”重试机器人权限；回复“取消”可放弃。此时尚未发起写入审批。`;
  try { await tell(caller.chatId, question); }
  catch (error) { await db.execute({ sql: "UPDATE permission_requests SET status = 'table_failed' WHERE id = ?", args: [id] }); throw error; }
  return { status: 'permission_choice_required', requestId: id, question };
}

export async function requestTableOperation(caller: Caller, intent: TableIntent) {
  if (intent.permissionMode === 'user') {
    const token = await userAccessToken(caller.userId, config.tableOAuthScopes);
    if (!token) return createPending(caller, intent, '当前用户令牌尚无表格权限');
    try { return await perform(caller, intent, 'user', token); }
    catch (error) {
      if (!authError(error)) throw error;
      if (String(error).includes('目标资源没有编辑权限')) throw error;
      return createPending(caller, intent, `用户表格权限不足：${String(error)}`);
    }
  }
  try { return await perform(caller, intent, 'app'); }
  catch (error) {
    if (!authError(error)) throw error;
    return createPending(caller, intent, `机器人应用无法访问目标表格：${String(error)}`);
  }
}

async function pendingFor(caller: Caller): Promise<Pending | undefined> {
  return (await (await stateDb()).execute({ sql: `SELECT * FROM permission_requests WHERE requester_id = ? AND chat_id = ? AND status IN ('table_choice','table_auth') AND expires_at > ? ORDER BY created_at DESC LIMIT 1`, args: [caller.userId, caller.chatId, Date.now()] })).rows[0] as unknown as Pending | undefined;
}

export async function supersedeTablePermissionForNewRequest(caller: Caller, content: string): Promise<boolean> {
  // A complete new request is a retry, even when an earlier permission question is still pending.
  // Keep short permission replies in the existing choice flow.
  if (!/https?:\/\/[^\s)]+\/(?:wiki|sheets|base)\/[\w@.-]+/i.test(content)
    || !/(?:写入|填写|更新|修改|追加|新增|创建|读取|查看|检查)/.test(content)) return false;
  await (await stateDb()).execute({
    sql: `UPDATE permission_requests SET status = 'superseded' WHERE requester_id = ? AND chat_id = ? AND status IN ('table_choice','table_auth','table_failed') RETURNING id`,
    args: [caller.userId, caller.chatId],
  });
  return true;
}

async function sendLink(caller: Caller, pending: Pending, intent: TableIntent) {
  const scopes = [...new Set([...CALENDAR_SCOPES, ...config.tableOAuthScopes])];
  let link: string;
  try { link = authorizationUrl(pending.id, scopes); }
  catch (error) {
    await tell(caller.chatId, `${String(error)}。当前没有发起写入审批；可回复“应用权限”重试或“取消”。`);
    return;
  }
  await (await stateDb()).execute({ sql: "UPDATE permission_requests SET status = 'table_auth', payload = ? WHERE id = ? AND status = 'table_choice'", args: [JSON.stringify({ ...intent, permissionMode: 'user' }), pending.id] });
  await tell(caller.chatId, `请以你自己的飞书账号打开并同意表格授权：${link}\n授权完成后，我会再发写入审批卡片；授权本身不会修改表格。`);
}

export async function resolveTablePermissionReply(caller: Caller, content: string): Promise<boolean> {
  const answer = content.trim().replace(/^@\S+\s*/, '');
  const db = await stateDb();
  let pending = await pendingFor(caller);
  if (!pending && /^(重新授权|重新申请授权|reauthorize)$/i.test(answer)) {
    pending = (await db.execute({ sql: `SELECT * FROM permission_requests WHERE requester_id = ? AND chat_id = ? AND status = 'table_failed' AND expires_at > ? ORDER BY created_at DESC LIMIT 1`, args: [caller.userId, caller.chatId, Date.now()] })).rows[0] as unknown as Pending | undefined;
    if (pending) {
      const intent = JSON.parse(pending.payload) as TableIntent;
      const result = await createPending(caller, intent, '重新申请表格权限');
      pending = { ...pending, id: result.requestId, status: 'table_choice' };
      await sendLink(caller, pending, intent);
      return true;
    }
  }
  if (!pending) return false;
  if (/^(取消|算了|cancel)$/i.test(answer)) {
    await db.execute({ sql: "UPDATE permission_requests SET status = 'cancelled' WHERE id = ?", args: [pending.id] });
    await tell(caller.chatId, '已取消这次表格请求。');
    return true;
  }
  const intent = JSON.parse(pending.payload) as TableIntent;
  if (/^(用户权限|用户身份|user)$/i.test(answer)) {
    const token = await userAccessToken(caller.userId, config.tableOAuthScopes);
    if (token) {
      try {
        const result = await perform(caller, intent, 'user', token);
        await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [pending.id] });
        if ('result' in result) await tell(caller.chatId, JSON.stringify(result.result).slice(0, 3500));
        return true;
      } catch (error) {
        if (String(error).includes('目标资源没有编辑权限')) {
          await tell(caller.chatId, '你当前没有该表格的编辑权限，请联系所有者授权；未发起写入审批。');
          return true;
        }
        if (!authError(error)) throw error;
      }
    }
    await sendLink(caller, pending, intent);
    return true;
  }
  if (/^(?:B|应用权限|应用身份|app)(?:[\s，,、：:].*)?$/i.test(answer)
    || /^已开通.{0,40}应用.{0,40}权限.{0,30}(?:重试)?$/i.test(answer)) {
    try {
      const result = await perform(caller, intent, 'app');
      await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [pending.id] });
      if ('result' in result) await tell(caller.chatId, JSON.stringify(result.result).slice(0, 3500));
    } catch (error) { await tell(caller.chatId, `应用权限仍无法访问：${String(error)}。请回复“用户权限”授权，或“取消”。`); }
    return true;
  }
  return false;
}

export async function completeTableAuthorization(state: string, code: string) {
  const db = await stateDb();
  const claimed = await db.execute({ sql: `UPDATE permission_requests SET status = 'table_auth_processing' WHERE id = ? AND status = 'table_auth' AND expires_at > ? RETURNING *`, args: [state, Date.now()] });
  const pending = claimed.rows[0] as unknown as Pending | undefined;
  if (!pending) throw new Error('表格授权请求已失效或已处理');
  const caller: Caller = { userId: pending.requester_id, chatId: pending.chat_id, chatType: pending.chat_type ?? 'p2p' };
  const intent = JSON.parse(pending.payload) as TableIntent;
  try {
    const token = await exchangeCodeForUser(code, caller.userId, [...new Set([...CALENDAR_SCOPES, ...config.tableOAuthScopes])]);
    const result = await perform(caller, intent, 'user', token);
    await db.execute({ sql: "UPDATE permission_requests SET status = 'completed' WHERE id = ?", args: [state] });
    if ('result' in result) await tell(caller.chatId, JSON.stringify(result.result).slice(0, 3500));
  } catch (error) {
    await db.execute({ sql: "UPDATE permission_requests SET status = 'table_failed' WHERE id = ?", args: [state] });
    const message = String(error).includes('目标资源没有编辑权限')
      ? '用户授权已完成，但你当前没有该表格的编辑权限。请联系所有者授权后重新发起表格请求；未发起写入审批。'
      : `表格授权或审批准备失败：${String(error)}。30 分钟内可回复“重新授权”重试。`;
    await tell(caller.chatId, message);
    throw error;
  }
}
