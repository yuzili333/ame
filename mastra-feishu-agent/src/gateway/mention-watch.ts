import type { LarkChannel, NormalizedMessage } from '@larksuiteoapi/node-sdk';
import type { OfficeApi } from '../office/client';
import { stateDb } from '../state/store';
import { createHash } from 'node:crypto';
import { EntityType, SpanType, type Span } from '@mastra/core/observability';
import { mastra } from '../mastra';
import { agent } from '../mastra/agents/agent';
import { historyText } from './chat-summary';

export function isOwnerMention(message: NormalizedMessage, ownerId?: string): boolean {
  return Boolean(ownerId && message.chatType === 'group' && !message.mentionAll && message.senderId !== ownerId
    && message.mentions.some(mention => mention.openId === ownerId));
}

async function sendToUser(channel: LarkChannel, ownerId: string, text: string, messageId: string): Promise<void> {
  const hex = createHash('sha256').update(`mention-backup:${messageId}:${ownerId}`).digest('hex');
  const uuid = `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
  const result = await channel.rawClient.request({ method: 'POST', url: '/open-apis/im/v1/messages', params: { receive_id_type: 'open_id' },
    data: { receive_id: ownerId, msg_type: 'text', content: JSON.stringify({ text }), uuid } });
  if (result.code !== 0) throw new Error(`飞书私聊备份发送失败：${result.code} ${result.msg}`);
}

export async function findGroupMembers(office: OfficeApi, chatId: string, ownerId: string, senderId: string): Promise<{ ownerPresent: boolean; senderName?: string }> {
  let token: string | undefined;
  let ownerPresent = false;
  let senderName: string | undefined;
  for (let page = 0; page < 100; page++) {
    const data = await office.call({ method: 'GET', url: `/open-apis/im/v1/chats/${encodeURIComponent(chatId)}/members`,
      params: { member_id_type: 'open_id', page_size: 100, ...(token ? { page_token: token } : {}) } });
    const items = (data.items ?? []) as Array<{ member_id?: string; name?: string }>;
    if (items.some(item => item.member_id === ownerId)) ownerPresent = true;
    senderName ??= items.find(item => item.member_id === senderId)?.name?.trim() || undefined;
    if (ownerPresent && senderName) return { ownerPresent, senderName };
    if (!data.has_more) return { ownerPresent, senderName };
    token = typeof data.page_token === 'string' ? data.page_token : undefined;
    if (!token) return { ownerPresent, senderName };
  }
  return { ownerPresent, senderName };
}

export function originalMessageLink(chatId: string, position?: string, apiLink?: string): string {
  const url = new URL('https://applink.co.cibfintech.com/client/chat/open');
  if (apiLink) {
    try {
      const supplied = new URL(apiLink);
      if (supplied.protocol === 'https:' && supplied.pathname === '/client/chat/open' && supplied.searchParams.get('openChatId') === chatId) {
        position ||= supplied.searchParams.get('position') ?? undefined;
      }
    } catch { /* Use the confirmed intranet AppLink. */ }
  }
  url.searchParams.set('openChatId', chatId);
  if (position && /^\d+$/.test(position)) url.searchParams.set('position', position);
  return url.toString();
}

export async function backupMention(message: NormalizedMessage, ownerId: string, office: OfficeApi, channel: LarkChannel): Promise<void> {
  if (!isOwnerMention(message, ownerId)) return;
  let span: Span<SpanType.GENERIC> | undefined;
  try {
    span = mastra.observability.getDefaultInstance()?.startSpan({
      type: SpanType.GENERIC,
      name: 'Feishu owner mention backup',
      entityType: EntityType.AGENT,
      entityId: agent.id,
      entityName: agent.name,
      input: { messageId: message.messageId, chatId: message.chatId, ownerId },
      metadata: { source: 'feishu-gateway', capability: 'owner-mention-backup' },
    });
  } catch (error) { console.error('[mention backup trace start]', error instanceof Error ? error.message : String(error)); }
  let status: 'sent' | 'skipped' | 'duplicate' | undefined;
  let failure: Error | undefined;
  try { status = await deliverMentionBackup(message, ownerId, office, channel); }
  catch (error) { failure = error instanceof Error ? error : new Error(String(error)); }
  try {
    if (failure) span?.error({ error: failure, endSpan: true });
    else span?.end({ output: { status } });
    await mastra.observability.flush();
  } catch (error) { console.error('[mention backup trace write]', error instanceof Error ? error.message : String(error)); }
  if (failure) throw failure;
}

async function deliverMentionBackup(message: NormalizedMessage, ownerId: string, office: OfficeApi, channel: LarkChannel): Promise<'sent' | 'skipped' | 'duplicate'> {
  const db = await stateDb();
  const inserted = await db.execute({ sql: "INSERT OR IGNORE INTO mention_backups(message_id,chat_id,owner_id,status,created_at) VALUES (?,?,?,'pending',?) RETURNING message_id",
    args: [message.messageId, message.chatId, ownerId, Date.now()] });
  if (!inserted.rows.length) {
    const resumed = await db.execute({ sql: "UPDATE mention_backups SET status='pending' WHERE message_id=? AND status='failed' RETURNING message_id", args: [message.messageId] });
    if (!resumed.rows.length) return 'duplicate';
  }
  try {
    const members = await findGroupMembers(office, message.chatId, ownerId, message.senderId);
    if (!members.ownerPresent) {
      await db.execute({ sql: "UPDATE mention_backups SET status='skipped' WHERE message_id=?", args: [message.messageId] });
      return 'skipped';
    }
    let name = message.chatId;
    try {
      const chat = await office.call({ method: 'GET', url: `/open-apis/im/v1/chats/${encodeURIComponent(message.chatId)}` });
      name = (chat.name as string | undefined) || message.chatId;
    } catch { /* Preserve the chat ID when group metadata is unavailable. */ }
    let original: { sender?: { sender_name?: string }; message_position?: string; message_app_link?: string } | undefined;
    try {
      const result = await office.call({ method: 'GET', url: `/open-apis/im/v1/messages/${encodeURIComponent(message.messageId)}`, params: { with_sender_name: true } });
      original = ((result.items as Array<typeof original> | undefined) ?? [])[0];
    } catch { /* The group link and event content remain usable. */ }
    let sender = members.senderName || original?.sender?.sender_name || message.senderName;
    if (!sender && message.senderId.startsWith('ou_')) {
      try {
        const person = await office.call({ method: 'GET', url: `/open-apis/contact/v3/users/${encodeURIComponent(message.senderId)}`, params: { user_id_type: 'open_id' } });
        sender = (person.user as { name?: string } | undefined)?.name;
      } catch { /* Preserve a readable ID fallback when contact permission is unavailable. */ }
    }
    sender ||= `姓名不可获取（${message.senderId}）`;
    const time = new Date(message.createTime + 8 * 60 * 60_000).toISOString().replace('T', ' ').slice(0, 19);
    const link = originalMessageLink(message.chatId, original?.message_position, original?.message_app_link);
    const linkLabel = new URL(link).searchParams.has('position') ? '查看原消息' : '打开群聊（按时间查找）';
    const text = `📌 群聊 @消息备份\n群聊：${name}\n发送者：${sender}\n时间：${time} 北京时间\n原文：${message.content || `[${message.rawContentType}消息]`}\n消息 ID：${message.messageId}\n${linkLabel}：${link}`;
    for (let attempt = 0; ; attempt++) {
      try { await sendToUser(channel, ownerId, text, message.messageId); break; }
      catch (error) {
        if (attempt >= 2) throw error;
        await new Promise(resolve => setTimeout(resolve, 500 * 2 ** attempt));
      }
    }
    await db.execute({ sql: "UPDATE mention_backups SET status='sent' WHERE message_id=?", args: [message.messageId] });
    return 'sent';
  } catch (error) {
    await db.execute({ sql: "UPDATE mention_backups SET status='failed' WHERE message_id=?", args: [message.messageId] });
    throw error;
  }
}

export async function recoverMentionBackups(office: OfficeApi, channel: LarkChannel): Promise<void> {
  const db = await stateDb();
  const rows = (await db.execute({ sql: "SELECT message_id, chat_id, owner_id FROM mention_backups WHERE status IN ('pending','failed') ORDER BY created_at LIMIT 100" })).rows;
  for (const row of rows) {
    const messageId = String(row.message_id), chatId = String(row.chat_id), ownerId = String(row.owner_id);
    try {
      const result = await office.call({ method: 'GET', url: `/open-apis/im/v1/messages/${encodeURIComponent(messageId)}`, params: { with_sender_name: true } });
      const item = ((result.items as Array<{ deleted?: boolean; msg_type?: string; create_time?: string; sender?: { id?: string; sender_name?: string }; body?: { content?: string }; mentions?: Array<{ key: string; id?: string; name?: string }> }> | undefined) ?? [])[0];
      if (!item || item.deleted) {
        await db.execute({ sql: "UPDATE mention_backups SET status='skipped' WHERE message_id=?", args: [messageId] });
        continue;
      }
      await db.execute({ sql: "UPDATE mention_backups SET status='failed' WHERE message_id=? AND status='pending'", args: [messageId] });
      await backupMention({ messageId, chatId, chatType: 'group', senderId: item.sender?.id ?? '', senderName: item.sender?.sender_name,
        content: historyText({ msg_type: item.msg_type, body: item.body }), rawContentType: item.msg_type ?? 'unknown', resources: [],
        mentions: (item.mentions ?? []).map(m => ({ key: m.key, openId: m.id, name: m.name })),
        mentionAll: (item.mentions ?? []).some(m => m.key === '@_all'), mentionedBot: false, createTime: Number(item.create_time) || Date.now() }, ownerId, office, channel);
    } catch (error) { console.error('[mention backup recovery]', error instanceof Error ? error.message : String(error)); }
  }
}
