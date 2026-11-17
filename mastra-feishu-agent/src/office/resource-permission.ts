import { randomUUID } from 'node:crypto';
import { isAuthorizedCaller, type Caller } from '../identity';
import { requestOfficeApproval } from '../runtime/approval-bridge';
import { getOfficeApi, sendOfficeText } from '../runtime/office-bridge';
import { getAction, isMessageRecalled, stateDb } from '../state/store';
import { actionSummary, type OfficeAction } from './actions';
import type { OfficeApi } from './client';
import { assertTaskTimeRange, reconcileTaskTime } from './task-time';
import { authorizationScopes, authorizationUrl, exchangeCodeForUser, userAccessToken, UserAuthorizationRequiredError } from './user-auth';
import { isResourceWrite, parseResourceIntent, prepareResourceWrite, readResource, requiredResourceScopes, resourceFamily, resourceMode, resourceReadSummary, ResourceAccessError, type ResourceIntent, type ResourceWrite } from './resource-operation';

type Pending = { id: string; requester_id: string; chat_id: string; chat_type: Caller['chatType']; payload: string; status: string; expires_at: number; origin_message_id: string | null };
type AuthorizationAction = Extract<OfficeAction, { kind: 'resource_authorize' }>;
export type ResourceContinuation = { originalMessage: string; messageCreatedAt: number };
type Continue = (caller: Caller, source: ResourceContinuation, result: Record<string, unknown>, requestId: string) => Promise<void>;
let continueRequest: Continue | undefined;
export function setResourceContinuation(value: Continue | undefined) { continueRequest = value; }
const ttl = 30 * 60_000;
const scopeError = (error: unknown) => error instanceof UserAuthorizationRequiredError || /Feishu API (401|99991663|99991664|99991668|99991671|99991672|99991676|99991677|99991679)\b/.test(String(error));
const accessError = (error: unknown) => error instanceof ResourceAccessError || /Feishu API (403|191002|193002|1470403)\b/.test(String(error));
function payload(row: Pending) { const data = JSON.parse(row.payload) as { intent: unknown; scopes: string[]; continuation?: ResourceContinuation }; return { intent: normalizeIntent(data.intent, data.continuation), scopes: data.scopes, continuation: data.continuation }; }
function normalizeIntent(input: unknown, continuation?: ResourceContinuation): ResourceIntent {
  const intent = parseResourceIntent(input);
  if (intent.kind !== 'task_create' && intent.kind !== 'task_update') return intent;
  const normalized = continuation ? reconcileTaskTime(intent, continuation.originalMessage, continuation.messageCreatedAt) : intent;
  assertTaskTimeRange(normalized);
  return normalized;
}
function callerFor(row: Pending): Caller { return { userId: row.requester_id, chatId: row.chat_id, chatType: row.chat_type, messageId: row.origin_message_id ?? undefined }; }
async function assertCaller(caller: Caller) {
  if (!isAuthorizedCaller(caller)) throw new Error('Requester is not allowed or direct chat is not allowlisted');
  if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，操作已取消');
}

async function perform(caller: Caller, intent: ResourceIntent, token?: string, guard: () => Promise<void> = async () => {}) {
  await assertCaller(caller);
  await guard();
  if (isResourceWrite(intent)) {
    const action = await prepareResourceWrite(getOfficeApi(), intent, token);
    await assertCaller(caller);
    await guard();
    const db = await stateDb();
    // A crash after the card was sent must not create a second write approval
    // when the same saved intent is resumed. Requests without a message ID
    // (Studio) are independent and cannot be deduplicated by origin.
    if (caller.messageId) {
      const previous = (await db.execute({ sql: `SELECT id,status,expires_at,result FROM actions WHERE requester_id=? AND chat_id=? AND origin_message_id=? AND kind=? AND payload=?
        AND status IN ('pending','executing','succeeded','failed','interrupted') ORDER BY created_at DESC,rowid DESC LIMIT 1`, args: [caller.userId, caller.chatId, caller.messageId, action.kind, JSON.stringify(action)] })).rows[0];
      const authorizationOnly = previous?.status === 'succeeded' && previous.result && (() => { try { return JSON.parse(String(previous.result)).status === 'authorization_required'; } catch { return false; } })();
      if (previous && !authorizationOnly) {
        if (previous.status === 'failed' || previous.status === 'interrupted') throw new Error('原操作结果需要人工核对，不能通过重新授权重复执行；请核对后发送新的请求');
        if (previous.status !== 'pending' || Number(previous.expires_at) > Date.now()) return { status: String(previous.status), approvalId: String(previous.id), expiresAt: Number(previous.expires_at) };
      }
    }
    const proposal = await requestOfficeApproval(caller, action);
    try { await guard(); }
    catch (error) {
      await db.execute({ sql: "UPDATE actions SET status='denied',result='授权续办已取消' WHERE id=? AND status='pending'", args: [proposal.approvalId] });
      throw error;
    }
    return proposal;
  }
  const result = await readResource(getOfficeApi(), intent, token);
  await assertCaller(caller);
  await guard();
  return { status: 'read', result };
}

async function requestAuthorization(caller: Caller, intent: ResourceIntent, continuation?: ResourceContinuation) {
  await assertCaller(caller);
  // Validate deployment configuration before issuing a card which cannot be used.
  authorizationUrl('configuration-check', requiredResourceScopes(intent));
  const db = await stateDb();
  const id = randomUUID();
  const scopes = await authorizationScopes(caller.userId, requiredResourceScopes(intent));
  const selected = { ...intent, permissionMode: 'user' };
  await db.execute({ sql: `INSERT INTO permission_requests(id,requester_id,chat_id,chat_type,payload,status,created_at,expires_at,origin_message_id)
    VALUES (?,?,?,?,?,'resource_approval',?,?,?)`, args: [id, caller.userId, caller.chatId, caller.chatType, JSON.stringify({ intent: selected, scopes, continuation }), Date.now(), Date.now() + ttl, caller.messageId ?? null] });
  try {
    const proposal = await requestOfficeApproval(caller, { kind: 'resource_authorize', requestId: id, resource: resourceFamily(intent), summary: isResourceWrite(intent) ? actionSummary({ ...intent, permissionMode: 'user' }) : resourceReadSummary(intent) });
    return { ...proposal, status: 'authorization_required', requestId: id };
  } catch (error) {
    await db.execute({ sql: "UPDATE permission_requests SET status='resource_failed' WHERE id=?", args: [id] });
    throw error;
  }
}

export async function requestResourceOperation(caller: Caller, input: ResourceIntent, continuation?: ResourceContinuation) {
  await assertCaller(caller);
  const intent = normalizeIntent(input, continuation);
  const mode = resourceMode(intent);
  try {
    const token = mode === 'user' ? await userAccessToken(caller.userId, requiredResourceScopes(intent)) : undefined;
    if (mode === 'user' && !token) return requestAuthorization(caller, intent, continuation);
    return await perform(caller, intent, token ?? undefined);
  } catch (error) {
    if (scopeError(error) || (mode === 'app' && accessError(error))) return requestAuthorization(caller, intent, continuation);
    if (accessError(error)) throw new ResourceAccessError('当前用户没有目标日历或任务的操作权限，请联系所有者授权后重新发起；OAuth 不能授予资源协作权限');
    throw error;
  }
}

export async function beginResourceAuthorization(action: AuthorizationAction, executionKey: string) {
  const approval = await getAction(executionKey);
  if (!approval || approval.kind !== 'resource_authorize' || approval.status !== 'executing' || JSON.parse(approval.payload).requestId !== action.requestId) throw new Error('必须先由原请求者确认授权审批');
  const db = await stateDb();
  const row = (await db.execute({ sql: "SELECT * FROM permission_requests WHERE id=? AND status='resource_approval' AND expires_at>?", args: [action.requestId, Date.now()] })).rows[0] as unknown as Pending | undefined;
  if (!row || row.requester_id !== approval.requester_id || row.chat_id !== approval.chat_id) throw new Error('授权请求已失效');
  await assertCaller(callerFor(row));
  const link = authorizationUrl(row.id, payload(row).scopes);
  const changed = await db.execute({ sql: "UPDATE permission_requests SET status='resource_auth' WHERE id=? AND status='resource_approval' RETURNING id", args: [row.id] });
  if (!changed.rows.length) throw new Error('授权请求已失效或已处理');
  return { authorizationUrl: link };
}

export async function executeResourceWrite(caller: Caller, action: ResourceWrite, executionKey: string, api: OfficeApi = getOfficeApi()) {
  await assertCaller(caller);
  const mode = resourceMode(action);
  try {
    const token = mode === 'user' ? await userAccessToken(caller.userId, requiredResourceScopes(action)) : undefined;
    if (mode === 'user' && !token) return requestAuthorization(caller, action);
    return await api.execute(action, executionKey, token ?? undefined, caller.userId);
  } catch (error) {
    // Only explicit permission rejections can start authorization. Timeouts or
    // other ambiguous write outcomes must remain failed and are never replayed.
    if (scopeError(error) || (mode === 'app' && accessError(error))) return requestAuthorization(caller, action);
    if (accessError(error)) throw new ResourceAccessError('当前用户没有目标资源的操作权限，请联系所有者授权；未自动重试');
    throw error;
  }
}

async function ensureResumable(row: Pending) {
  await assertCaller(callerFor(row));
  const current = (await (await stateDb()).execute({ sql: 'SELECT status,expires_at FROM permission_requests WHERE id=?', args: [row.id] })).rows[0];
  if (current?.status !== 'resource_auth_processing' || Number(current.expires_at) <= Date.now()) throw new Error('授权请求已失效或已取消');
}

export async function completeResourceAuthorization(state: string, code: string) {
  const db = await stateDb();
  const claimed = await db.execute({ sql: "UPDATE permission_requests SET status='resource_auth_processing' WHERE id=? AND status='resource_auth' AND expires_at>? RETURNING *", args: [state, Date.now()] });
  const row = claimed.rows[0] as unknown as Pending | undefined;
  if (!row) throw new Error('授权请求已失效或已处理');
  const caller = callerFor(row);
  try {
    await ensureResumable(row);
    const { intent, scopes, continuation } = payload(row);
    await exchangeCodeForUser(code, caller.userId, scopes);
    await ensureResumable(row);
    const token = await userAccessToken(caller.userId, requiredResourceScopes(intent));
    if (!token) throw new UserAuthorizationRequiredError('本次授权未包含操作所需范围，请管理员开通并发布权限后重新授权');
    const result = await perform(caller, intent, token, () => ensureResumable(row));
    const finished = await db.execute({ sql: "UPDATE permission_requests SET status='completed' WHERE id=? AND status='resource_auth_processing' RETURNING id", args: [state] });
    if (!finished.rows.length) {
      if ('approvalId' in result) await db.execute({ sql: "UPDATE actions SET status='denied' WHERE id=? AND status='pending'", args: [result.approvalId] });
      throw new Error('授权请求已取消');
    }
    if ('result' in result) {
      if (continuation && continueRequest && caller.messageId) {
        await continueRequest(caller, continuation, result.result, state);
        return;
      }
      // Keep pagination metadata even when a response is too large for one message.
      const text = JSON.stringify(result.result, null, 2);
      const suffix = text.length > 10_000 ? `\n结果已截断；has_more=${String(result.result.has_more ?? false)}，page_token=${String(result.result.page_token ?? '')}；请缩小范围或继续分页。` : '';
      await sendOfficeText(caller.chatId, `${isResourceWrite(intent) ? '操作' : resourceReadSummary(intent)}结果：\n${text.slice(0, 10_000)}${suffix}`);
    }
  } catch (error) {
    await db.execute({ sql: "UPDATE permission_requests SET status='resource_failed' WHERE id=? AND status='resource_auth_processing'", args: [state] });
    const current = (await db.execute({ sql: 'SELECT status FROM permission_requests WHERE id=?', args: [state] })).rows[0];
    if (current?.status !== 'cancelled' && !(await isMessageRecalled(caller.messageId))) await sendOfficeText(caller.chatId, accessError(error)
      ? '用户授权已完成，但你没有目标日历或任务的操作权限。请联系所有者授权后重新发起；未执行写入。'
      : `授权续办失败：${error instanceof Error ? error.message : String(error)}。可回复“重新授权”申请新的审批；原操作不会自动重复执行。`);
    throw error;
  }
}

export async function cancelResourceAuthorization(state: string) {
  const db = await stateDb();
  const changed = await db.execute({ sql: "UPDATE permission_requests SET status='cancelled' WHERE id=? AND status IN ('resource_approval','resource_auth','resource_auth_processing') RETURNING *", args: [state] });
  const row = changed.rows[0] as unknown as Pending | undefined;
  if (!row) return;
  await db.execute({ sql: "UPDATE actions SET status='denied' WHERE kind='resource_authorize' AND status='pending' AND json_extract(payload,'$.requestId')=?", args: [state] });
  await sendOfficeText(row.chat_id, '已取消这次日历/任务授权请求，未执行操作。');
}

export async function resolveResourcePermissionReply(caller: Caller, text: string): Promise<boolean> {
  if (!/^(取消|算了|cancel|重新授权|重新申请授权|reauthorize)$/i.test(text.trim())) return false;
  await assertCaller(caller);
  const db = await stateDb();
  const row = (await db.execute({ sql: `SELECT * FROM permission_requests WHERE requester_id=? AND chat_id=? AND status IN ('resource_approval','resource_auth','resource_auth_processing','resource_failed') AND expires_at>? ORDER BY created_at DESC,rowid DESC LIMIT 1`, args: [caller.userId, caller.chatId, Date.now()] })).rows[0] as unknown as Pending | undefined;
  if (!row) return false;
  if (/^(取消|算了|cancel)$/i.test(text.trim())) {
    if (row.status === 'resource_failed') await db.execute({ sql: "UPDATE permission_requests SET status='cancelled' WHERE id=?", args: [row.id] });
    else await cancelResourceAuthorization(row.id);
  } else {
    if (row.status !== 'resource_failed') { await sendOfficeText(caller.chatId, '已有授权请求正在等待处理，请完成当前授权或先取消。'); return true; }
    await assertCaller(callerFor(row));
    await requestAuthorization(callerFor(row), payload(row).intent, payload(row).continuation);
    await db.execute({ sql: "UPDATE permission_requests SET status='superseded' WHERE id=?", args: [row.id] });
  }
  return true;
}
