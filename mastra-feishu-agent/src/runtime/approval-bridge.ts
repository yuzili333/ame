import { isAuthorizedCaller, type Caller } from '../identity';
import { actionSummary, type OfficeAction } from '../office/actions';
import { proposeAction } from '../state/approvals';
import { isMessageRecalled, stateDb } from '../state/store';
import { sendStudioApprovalCard } from './office-bridge';
import { config } from '../config';

type Notifier = (chatId: string, actionId: string, summary: string) => Promise<void>;
let notifier: Notifier | undefined;

export function setApprovalNotifier(value: Notifier | undefined) { notifier = value; }

export async function requestOfficeApproval(caller: Caller, action: OfficeAction) {
  if (!isAuthorizedCaller(caller)) {
    throw new Error('Requester is not allowed or direct chat is not allowlisted');
  }
  if (action.kind === 'message_send' && !config.allowedUsers.has(caller.userId)) {
    throw new Error('Cross-chat message sending requires an allowlisted user');
  }
  const proposal = await proposeAction(caller, action);
  try {
    if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，本轮操作已取消');
    await (notifier ?? sendStudioApprovalCard)(caller.chatId, proposal.id, actionSummary(action));
    if (await isMessageRecalled(caller.messageId)) throw new Error('原始消息已撤回，本轮操作已取消');
  } catch (error) {
    const db = await stateDb();
    await db.execute({ sql: "UPDATE actions SET status = 'failed', result = ? WHERE id = ? AND status = 'pending'", args: [String(error), proposal.id] });
    throw error;
  }
  return { status: 'pending', approvalId: proposal.id, expiresAt: proposal.expires_at };
}
