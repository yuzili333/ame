import { randomUUID } from 'node:crypto';
import type { Client } from '@libsql/client';
import { officeAction, type OfficeAction } from '../office/actions';
import type { Caller } from '../identity';
import { getAction, isMessageRecalled, stateDb, type StoredAction } from './store';

export const APPROVAL_TTL_MS = 30 * 60 * 1000;

export type ProposedAction = OfficeAction | { kind: 'mastra_tool'; toolName: string; arguments: unknown };

export async function proposeAction(caller: Caller, action: ProposedAction, options?: { toolCallId?: string; sessionId?: string }): Promise<StoredAction> {
  if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，本轮操作已取消');
  const db = await stateDb();
  const now = Date.now();
  const id = randomUUID();
  await db.execute({
    sql: `INSERT INTO actions(id, kind, payload, requester_id, chat_id, chat_type, status, created_at, expires_at, tool_call_id, session_id, origin_message_id)
          VALUES (?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?, ?, ?)`,
    args: [id, action.kind, JSON.stringify(action), caller.userId, caller.chatId, caller.chatType, now, now + APPROVAL_TTL_MS, options?.toolCallId ?? null, options?.sessionId ?? null, caller.messageId ?? null],
  });
  if (await isMessageRecalled(caller.messageId)) {
    await db.execute({ sql: "UPDATE actions SET status = 'denied', result = '原始消息已撤回' WHERE id = ? AND status = 'pending'", args: [id] });
    throw new Error('原始消息已撤回，本轮操作已取消');
  }
  return (await getAction(id))!;
}

async function transition(db: Client, action: StoredAction, status: StoredAction['status'], result?: string) {
  await db.execute({ sql: 'UPDATE actions SET status = ?, result = ? WHERE id = ?', args: [status, result ?? null, action.id] });
}

export type Decision = { status: StoredAction['status']; result?: unknown; action?: StoredAction };

export async function expireAction(id: string): Promise<StoredAction | null> {
  const db = await stateDb();
  const claimed = await db.execute({
    sql: "UPDATE actions SET status = 'expired' WHERE id = ? AND status = 'pending' AND expires_at <= ? RETURNING id",
    args: [id, Date.now()],
  });
  return claimed.rows.length ? getAction(id) : null;
}

export async function decideAction(
  id: string,
  caller: Pick<Caller, 'userId' | 'chatId'>,
  approved: boolean,
  executeOffice: (action: OfficeAction, executionKey: string) => Promise<unknown>,
  resumeTool?: (action: StoredAction, approved: boolean) => Promise<void>,
): Promise<Decision> {
  const db = await stateDb();
  const action = await getAction(id);
  if (!action || action.requester_id !== caller.userId || action.chat_id !== caller.chatId) return { status: 'denied' };
  if (await isMessageRecalled(action.origin_message_id ?? undefined)) {
    await db.execute({ sql: "UPDATE actions SET status = 'denied', result = '原始消息已撤回' WHERE id = ? AND status = 'pending'", args: [id] });
    return { status: 'denied', action };
  }
  if (action.status !== 'pending') return { status: action.status, action };
  if (action.expires_at < Date.now()) {
    await db.execute({ sql: "UPDATE actions SET status = 'expired' WHERE id = ? AND status = 'pending'", args: [id] });
    return { status: 'expired', action };
  }
  const next = approved ? 'executing' : 'denied';
  const claimed = await db.execute({
    sql: `UPDATE actions SET status = ? WHERE id = ? AND status = 'pending' RETURNING id`,
    args: [next, id],
  });
  if (!claimed.rows.length) return { status: (await getAction(id))?.status ?? 'denied' };
  if (!approved) {
    if (action.kind === 'mastra_tool' && resumeTool) await resumeTool(action, false);
    return { status: 'denied', action };
  }
  try {
    if (action.kind === 'mastra_tool') {
      if (!resumeTool) throw new Error('The original tool run is no longer available');
      await resumeTool(action, true);
      await transition(db, action, 'succeeded');
      return { status: 'succeeded', action };
    }
    const result = await executeOffice(officeAction.parse(JSON.parse(action.payload)), action.id);
    await transition(db, action, 'succeeded', JSON.stringify(result));
    return { status: 'succeeded', result, action };
  } catch (error) {
    // A failed external call may have committed remotely. Never replay it automatically.
    await transition(db, action, 'failed', error instanceof Error ? error.message : String(error));
    return { status: 'failed', result: error instanceof Error ? error.message : String(error), action };
  }
}
