import type { Caller } from '../identity';
import { isMessageRecalled, stateDb } from '../state/store';
import { schedulingText } from './scheduling-text';

// Only an explicit change of resource type inherits a recent user-authored
// schedule. Never use assistant history as a source of dates or timestamps.
export async function resolveScheduleSource(caller: Caller, originalMessage: string, messageCreatedAt: number) {
  const source = { originalMessage, messageCreatedAt, needsTime: false };
  const scheduling = schedulingText(originalMessage);
  const changesType = /(?:改为|改成|更改为|换成)\s*(?:创建|新增)?\s*[“"「]?(?:一个)?(?:日程|任务)/.test(scheduling);
  const suppliesTime = /今天|明天|后天|\d{1,2}[:：]\d{2}|\d{1,2}点|\d{1,2}月|\d{4}[-/.]|(?:周|星期)[一二三四五六日天]/.test(scheduling);
  const db = await stateDb();
  if (changesType && !suppliesTime) {
    const row = (await db.execute({ sql: 'SELECT * FROM schedule_sources WHERE requester_id=? AND chat_id=?', args: [caller.userId, caller.chatId] })).rows[0];
    const age = row ? messageCreatedAt - Number(row.message_created_at) : Infinity;
    if (!row || age < 0 || age > 30 * 60_000 || await isMessageRecalled(String(row.message_id))) return { ...source, needsTime: true };
    return { originalMessage: `原请求：${String(row.original_message)}\n本轮更正：${originalMessage}`, messageCreatedAt: Number(row.message_created_at), needsTime: false };
  }
  if (suppliesTime && /任务|日程/.test(scheduling) && /创建|新增|改为|改成|更改为|换成/.test(scheduling) && caller.messageId) {
    await db.execute({ sql: `INSERT INTO schedule_sources(requester_id,chat_id,message_id,original_message,message_created_at) VALUES (?,?,?,?,?)
      ON CONFLICT(requester_id,chat_id) DO UPDATE SET message_id=excluded.message_id,original_message=excluded.original_message,message_created_at=excluded.message_created_at`,
      args: [caller.userId, caller.chatId, caller.messageId, originalMessage, messageCreatedAt] });
  }
  return source;
}
