import { createLarkChannel, type LarkChannel } from '@larksuiteoapi/node-sdk';
import { requireFeishuConfig } from '../config';
import { approvalCard } from '../gateway/cards';
import { safeSdkLogger } from '../gateway/safe-sdk-logger';
import { OfficeApi, createOfficeTransport } from '../office/client';

let current: OfficeApi | undefined;
let studioChannel: LarkChannel | undefined;
export function setOfficeApi(api: OfficeApi | undefined) { current = api; }
function getStudioChannel(): LarkChannel {
  studioChannel ??= createLarkChannel({ ...requireFeishuConfig(), transport: 'websocket', logger: safeSdkLogger });
  return studioChannel;
}
export function getOfficeApi(): OfficeApi {
  return current ?? new OfficeApi(createOfficeTransport(getStudioChannel()));
}

export async function sendStudioApprovalCard(chatId: string, actionId: string, summary: string) {
  await getStudioChannel().send(chatId, { card: approvalCard(actionId, summary) });
}

let currentTextSender: ((chatId: string, text: string) => Promise<void>) | undefined;
export function setOfficeTextSender(sender: typeof currentTextSender) { currentTextSender = sender; }
export async function sendOfficeText(chatId: string, text: string) {
  if (currentTextSender) await currentTextSender(chatId, text);
  else await getStudioChannel().send(chatId, { markdown: text });
}
