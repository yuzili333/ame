import type { NormalizedMessage } from '@larksuiteoapi/node-sdk';
import { config } from './config';

export type Caller = { userId: string; chatId: string; chatType: 'p2p' | 'group'; messageId?: string };

export function isAuthorizedCaller(caller: Caller): boolean {
  if (!caller.userId || !caller.chatId) return false;
  return caller.chatType === 'group'
    || (config.allowedUsers.has(caller.userId) && config.allowedChats.has(caller.chatId));
}

export function isAuthorizedMessage(message: Pick<NormalizedMessage, 'senderId' | 'chatId' | 'chatType' | 'mentionedBot'>): boolean {
  return isAuthorizedCaller({ userId: message.senderId, chatId: message.chatId, chatType: message.chatType })
    && (message.chatType === 'p2p' || message.mentionedBot);
}

export function memoryKeys(caller: Caller) {
  const key = caller.chatType === 'group' ? `group:${caller.chatId}` : `dm:${caller.userId}`;
  return { sessionId: key, resourceId: key, threadId: key };
}
