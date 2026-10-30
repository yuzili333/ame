import { createLarkChannel, type CardActionEvent, type LarkChannel, type NormalizedMessage } from '@larksuiteoapi/node-sdk';
import { RequestContext } from '@mastra/core/request-context';
import type { Server } from 'node:http';
import { z } from 'zod';
import { config, requireFeishuConfig } from '../config';
import { isAuthorizedCaller, isAuthorizedMessage, memoryKeys, type Caller } from '../identity';
import { controller } from '../mastra';
import { minutesToTasksWorkflow } from '../mastra/workflows/minutes-to-tasks';
import { releaseMaterialsWorkflow } from '../mastra/workflows/release-materials';
import { officeTools, requestReleaseMaterials } from '../mastra/tools/office-tools';
import { OfficeApi, createOfficeTransport } from '../office/client';
import { ReleaseMaterialsService, releaseInput, type ReleaseManifest } from '../office/release-materials';
import { actionResultDetail, officeAction } from '../office/actions';
import { completeCalendarAuthorization, latestPermissionRequestId, resolvePermissionReply, setPermissionNotifier } from '../office/permission-choice';
import { startOAuthCallback, userAccessToken } from '../office/user-auth';
import { completeTableAuthorization, requestTableOperation, resolveTablePermissionReply, setTablePermissionNotifier, supersedeTablePermissionForNewRequest } from '../office/table-permission';
import { parseFirstEmptySheetCellRequest } from '../office/table-request';
import { setApprovalNotifier } from '../runtime/approval-bridge';
import { setOfficeApi, setOfficeTextSender } from '../runtime/office-bridge';
import { APPROVAL_TTL_MS, decideAction, expireAction, proposeAction } from '../state/approvals';
import { getAction, markMessageSeen, recoverInterruptedApprovals, recoverInterruptedWorkflows, stateDb } from '../state/store';
import { approvalCard, cardDecision, resolvedCard } from './cards';
import { safeSdkLogger } from './safe-sdk-logger';

type Session = Awaited<ReturnType<typeof controller.createSession>>;
const minutesCommand = /^\/minutes\s+([\w@.-]+)\s*$/i;
const idCommand = /^\/id$/i;
const releaseRetryCommand = /^重新发起(?:下发材料)?[。.!！]?$/;
const releaseStatusFollowup = /^(?:确认|已确认|批准|已批准|已点击确认|已点击批准|通过(?:新批次|这批|该批|下发材料)?(?:的)?(?:审批|确认)?卡片|(?:查看|核对|检查)(?:新批次|这批|该批|下发材料)?(?:的)?(?:审批|执行)?(?:结果|状态))[。.!！?？]?$/;
const officeToolNames = new Set(Object.keys(officeTools));

export class FeishuGateway {
  readonly channel: LarkChannel;
  readonly office: OfficeApi;
  private readonly sessions = new Map<string, Session>();
  private readonly queues = new Map<string, Promise<void>>();
  private readonly activeRuns = new Map<string, { unsubscribe: () => void }>();
  private oauthServer: Server | undefined;

  constructor(channel?: LarkChannel) {
    this.channel = channel ?? createLarkChannel({
      ...requireFeishuConfig(),
      transport: 'websocket',
      // Receive @bot messages from groups; the gateway checks their identity and mention.
      policy: { requireMention: true, dmMode: 'open' },
      safety: { chatQueue: { enabled: false }, batch: { text: { delayMs: 0 }, media: { delayMs: 0 } } },
      handshakeTimeoutMs: 15_000,
      logger: safeSdkLogger,
    });
    this.office = new OfficeApi(createOfficeTransport(this.channel));
  }

  async start() {
    await stateDb();
    await recoverInterruptedApprovals();
    await recoverInterruptedWorkflows();
    await controller.init();
    setOfficeApi(this.office);
    setOfficeTextSender(async (chatId, content) => { await this.channel.send(chatId, { markdown: content }); });
    setApprovalNotifier(async (chatId, actionId, summary) => {
      await this.channel.send(chatId, { card: approvalCard(actionId, summary) });
    });
    setPermissionNotifier(async (chatId, text) => { await this.channel.send(chatId, { text }); });
    setTablePermissionNotifier(async (chatId, text) => { await this.channel.send(chatId, { text }); });
    this.oauthServer = await startOAuthCallback(async (state, code) => {
      const row = (await (await stateDb()).execute({ sql: 'SELECT status FROM permission_requests WHERE id = ?', args: [state] })).rows[0];
      if (row?.status === 'table_auth') await completeTableAuthorization(state, code);
      else await completeCalendarAuthorization(state, code);
    });
    this.channel.on('message', message => { void this.enqueue(message).catch(error => this.report(error)); });
    this.channel.on('cardAction', event => { void this.onCardAction(event).catch(error => this.report(error)); });
    this.channel.on('error', error => this.report(error));
    await this.channel.connect();
  }

  async stop() {
    for (const run of this.activeRuns.values()) run.unsubscribe();
    this.activeRuns.clear();
    setApprovalNotifier(undefined);
    setPermissionNotifier(undefined);
    setTablePermissionNotifier(undefined);
    setOfficeApi(undefined);
    setOfficeTextSender(undefined);
    await new Promise<void>(resolve => this.oauthServer?.close(() => resolve()) ?? resolve());
    await this.channel.disconnect();
  }

  private report(error: unknown) {
    console.error(error instanceof Error ? error.message : String(error));
  }

  private enqueue(message: NormalizedMessage): Promise<void> {
    const before = this.queues.get(message.chatId) ?? Promise.resolve();
    const after = before.catch(() => {}).then(() => this.onMessage(message));
    this.queues.set(message.chatId, after);
    void after.finally(() => { if (this.queues.get(message.chatId) === after) this.queues.delete(message.chatId); }).catch(() => {});
    return after;
  }

  async onMessage(message: NormalizedMessage) {
    const content = message.content.trim();
    if (!content) return;
    const caller: Caller = { userId: message.senderId, chatId: message.chatId, chatType: message.chatType };
    const lookupId = idCommand.test(content) && (message.chatType === 'p2p' || message.mentionedBot);
    if (!lookupId && !isAuthorizedMessage(message)) return;
    const releaseRetry = releaseRetryCommand.test(content) ? await this.latestReleaseJob(caller) : null;
    if (!(await markMessageSeen(message.messageId))) return;
    try {
      if (lookupId) {
        const text = message.chatType === 'p2p'
          ? message.senderId.startsWith('ou_')
            ? `你的 open_id：${message.senderId}\n当前私聊 chat_id：${message.chatId}`
            : `当前私聊 chat_id：${message.chatId}\n飞书事件未提供 open_id，请检查机器人消息权限。`
          : `当前群 chat_id：${message.chatId}\n个人 open_id 请私聊机器人发送 /id。`;
        await this.channel.send(message.chatId, { text }, { replyTo: message.messageId });
        return;
      }
      if (releaseRetry) {
        await this.retryRelease(caller, releaseRetry, message.messageId, message.createTime);
        return;
      }
      if (releaseRetryCommand.test(content)) {
        await this.channel.send(caller.chatId, { text: '没有属于你的可重试下发批次；请提供完整的下发编号、两个文件夹链接和字段值。' }, { replyTo: message.messageId });
        return;
      }
      if (releaseStatusFollowup.test(content)) {
        const latest = await this.latestReleaseJob(caller);
        if (latest && Date.now() - latest.createdAt < APPROVAL_TTL_MS) {
          await this.channel.send(caller.chatId, { markdown: this.releaseJobMessage(latest) }, { replyTo: message.messageId });
          return;
        }
      }
      if (await this.handleKnownReleaseRequest(content, caller, message.messageId, message.createTime)) return;
      const freshTableRetry = await supersedeTablePermissionForNewRequest(caller, content);
      if (await resolveTablePermissionReply(caller, content)) return;
      if (await resolvePermissionReply(caller, content)) return;
      const tableRequest = parseFirstEmptySheetCellRequest(content);
      if (tableRequest) {
        await requestTableOperation(caller, tableRequest);
        return;
      }
      const command = minutesCommand.exec(content);
      if (command) {
        await this.runMinutesWorkflow(command[1], caller, message.messageId);
      } else {
        await this.runAgent(content, caller, message.messageId, message.createTime, freshTableRetry);
      }
    } catch (error) {
      this.report(error);
      await this.channel.send(message.chatId, { text: '处理失败。请检查模型代理、飞书权限和服务日志。' }, { replyTo: message.messageId });
    }
  }

  private async latestReleaseJob(caller: Caller): Promise<{ id: string; status: string; manifest: ReleaseManifest; approvalId?: string; createdAt: number; folderUrl?: string; result?: { folderUrl: string; files: Array<{ name: string; url: string }> } } | null> {
    if (!caller.userId || !caller.chatId) return null;
    const row = (await (await stateDb()).execute({
      sql: 'SELECT id, status, manifest_json, approval_id, created_at, folder_url, result_json FROM release_jobs WHERE requester_id = ? AND chat_id = ? ORDER BY created_at DESC LIMIT 1',
      args: [caller.userId, caller.chatId],
    })).rows[0];
    if (!row) return null;
    let status = String(row.status);
    const approvalId = typeof row.approval_id === 'string' ? row.approval_id : undefined;
    if (status === 'pending' && approvalId) {
      const action = await expireAction(approvalId) ?? await getAction(approvalId);
      if (action && ['failed', 'denied', 'expired'].includes(action.status)) {
        status = action.status;
        await (await stateDb()).execute({ sql: "UPDATE release_jobs SET status = ?, updated_at = ? WHERE approval_id = ? AND status = 'pending'", args: [status, Date.now(), approvalId] });
      }
    }
    return { id: String(row.id), status, manifest: JSON.parse(String(row.manifest_json)) as ReleaseManifest, approvalId,
      createdAt: Number(row.created_at), folderUrl: typeof row.folder_url === 'string' ? row.folder_url : undefined,
      result: typeof row.result_json === 'string' ? JSON.parse(row.result_json) as { folderUrl: string; files: Array<{ name: string; url: string }> } : undefined };
  }

  private releaseJobMessage(job: NonNullable<Awaited<ReturnType<FeishuGateway['latestReleaseJob']>>>): string {
    const id = job.manifest.input.releaseId;
    if (job.status === 'completed') {
      const folderUrl = job.result?.folderUrl ?? job.folderUrl;
      const files = job.result?.files ?? [];
      return `下发编号 ${id} 已执行完成。${folderUrl ? `\n材料文件夹：${folderUrl}` : ''}${files.length ? `\n文件：\n${files.map(file => `- ${file.name}：${file.url}`).join('\n')}` : ''}`;
    }
    if (job.status === 'pending') return `下发编号 ${id} 的整批确认卡片已发送${job.approvalId ? `（审批 ID：${job.approvalId}）` : ''}。当前状态：等待原请求者在卡片中确认。`;
    if (job.status === 'running' || job.status === 'creating_folder') return `下发编号 ${id} 正在执行，请稍后查看确认卡片的最终状态。`;
    if (job.status === 'needs_review') return `下发编号 ${id} 的执行结果需要人工核对；材料文件夹：${job.folderUrl ?? '尚未确认'}。系统不会自动重复写入。`;
    const label = job.status === 'failed' ? '失败' : job.status === 'denied' ? '被拒绝' : job.status === 'expired' ? '过期' : job.status;
    return `下发编号 ${id}：此前批次已${label}，旧审批卡片不再有效。请重新发起并核对新卡片。`;
  }

  private async handleKnownReleaseRequest(content: string, caller: Caller, replyTo: string, createdAt: number): Promise<boolean> {
    const latest = await this.latestReleaseJob(caller);
    if (!latest) return false;
    const input = latest.manifest.input;
    const values = [...new Set(Object.values(input.fields))].filter(Boolean);
    if (!content.includes(input.releaseId) || !content.includes(input.templateFolderUrl)
      || !content.includes(input.materialsParentUrl) || !values.every(value => content.includes(value))) return false;
    if (latest.status === 'pending') {
      await this.channel.send(caller.chatId, { text: `下发编号 ${input.releaseId} 已发送真正的整批确认卡片（审批 ID：${latest.approvalId ?? '待记录'}），仍在等待确认；本次没有再次创建审批。` }, { replyTo });
      return true;
    }
    if (latest.status === 'needs_review') {
      await this.channel.send(caller.chatId, { text: `下发编号 ${input.releaseId} 的上次执行结果需要人工核对；未重复创建目录或审批。` }, { replyTo });
      return true;
    }
    if (['failed', 'denied', 'expired'].includes(latest.status)) {
      try { await requestReleaseMaterials(input, caller, createdAt); }
      catch (error) {
        this.report(error);
        await this.channel.send(caller.chatId, { text: `重新预检下发材料失败：${error instanceof Error ? error.message : String(error)}。未发起写入。` }, { replyTo });
      }
      return true;
    }
    if (latest.status !== 'completed') return false;
    try {
      const files = await new ReleaseMaterialsService(this.office).listFolder(latest.manifest.materialsParentToken);
      const existing = files.find(file => file.type === 'folder' && file.name === input.releaseId);
      if (existing) {
        const url = existing.url ?? `${new URL(input.materialsParentUrl).origin}/drive/folder/${existing.token}`;
        await this.channel.send(caller.chatId, { text: `下发编号 ${input.releaseId} 的材料目录已存在：${url}。本次没有重复创建；如需新批次，请使用新的下发编号。` }, { replyTo });
        return true;
      }
      await (await stateDb()).execute({
        sql: "UPDATE release_jobs SET status = 'failed', error = ?, updated_at = ? WHERE id = ? AND status = 'completed'",
        args: ['远端材料目录已不存在；原完成记录失效，允许重新预检', Date.now(), latest.id],
      });
      await requestReleaseMaterials(input, caller, createdAt);
      return true;
    } catch (error) {
      this.report(error);
      await this.channel.send(caller.chatId, { text: `无法核验或重新预检材料目录：${error instanceof Error ? error.message : String(error)}。未发起写入。` }, { replyTo });
      return true;
    }
  }

  private async retryRelease(caller: Caller, latest: { status: string; manifest: ReleaseManifest; approvalId?: string }, replyTo: string, createdAt: number): Promise<void> {
    if (['failed', 'denied', 'expired'].includes(latest.status)) {
      const input = releaseInput.parse(latest.manifest.input);
      try {
        const result = await requestReleaseMaterials(input, caller, createdAt);
        if (result.status !== 'pending') await this.channel.send(caller.chatId, { text: `下发材料重试状态：${result.status}。` }, { replyTo });
      } catch (error) {
        this.report(error);
        await this.channel.send(caller.chatId, { text: `未能重新发起下发材料：${error instanceof Error ? error.message : String(error)}` }, { replyTo });
      }
      return;
    }
    const detail = latest.status === 'pending'
      ? `当前批次仍在等待确认${latest.approvalId ? `（审批 ID：${latest.approvalId}）` : ''}；如卡片已失效，请明确取消或等待其过期后再重试。`
      : latest.status === 'completed' ? '这批下发材料已完成，请提供新下发编号以创建另一批。'
        : '上次下发的执行结果需要人工核对；为避免重复创建，系统不会自动重试。';
    await this.channel.send(caller.chatId, { text: detail }, { replyTo });
  }

  private async sessionFor(caller: Caller) {
    const keys = memoryKeys(caller);
    let session = this.sessions.get(keys.sessionId);
    if (!session) {
      session = await controller.createSession({ id: keys.sessionId, resourceId: keys.resourceId, threadId: keys.threadId });
      await session.permissions.setForCategory({ category: 'execute', policy: 'ask' });
      this.sessions.set(keys.sessionId, session);
    }
    return session;
  }

  private async runAgent(content: string, caller: Caller, replyTo: string, messageCreatedAt = Date.now(), freshTableRetry = false) {
    const session = await this.sessionFor(caller);
    const permissionRequestBefore = await latestPermissionRequestId(caller);
    const sessionId = memoryKeys(caller).sessionId;
    if (this.activeRuns.has(sessionId)) {
      await this.channel.send(caller.chatId, { text: '上一轮操作仍在处理或等待确认，请先处理审批卡片。' }, { replyTo });
      return;
    }
    let reply = '';
    let finished = false;
    const finish = async (sendReply: boolean) => {
      if (finished) return;
      finished = true;
      unsubscribe();
      try {
        if (sendReply && reply.trim() && await latestPermissionRequestId(caller) === permissionRequestBefore) {
          let outgoing = reply;
          const latestRelease = await this.latestReleaseJob(caller);
          if (latestRelease && content.includes(latestRelease.manifest.input.releaseId) && /下发/.test(content)) {
            outgoing = this.releaseJobMessage(latestRelease);
          } else if (latestRelease && content.includes(latestRelease.manifest.input.releaseId)
            && ['failed', 'denied', 'expired'].includes(latestRelease.status)
            && latestRelease.approvalId && outgoing.includes(latestRelease.approvalId)
            && /待审批|等你审批|等待.*审批/.test(outgoing)) {
            const state = latestRelease.status === 'failed' ? '失败' : latestRelease.status === 'denied' ? '被拒绝' : '已过期';
            outgoing = `此前批次已${state}，旧审批卡片不再有效，也没有新的待审批批次。${caller.chatType === 'group' ? '请在群里发送“@机器人 重新发起”' : '请回复“重新发起”'}，我会重新预检并在有权限时发送新卡片。`;
          }
          await this.channel.send(caller.chatId, { markdown: outgoing }, { replyTo });
        }
      } finally {
        this.activeRuns.delete(sessionId);
      }
    };
    const unsubscribe = session.subscribe(event => {
      if (event.type === 'message_update' && event.event.type === 'text-delta') reply += event.event.delta;
      if (event.type === 'tool_approval_required') {
        void this.onToolApproval(event, caller, session).catch(error => this.report(error));
      }
      if (event.type === 'agent_end' && event.reason !== 'suspended') {
        void finish(event.reason !== 'error').catch(error => this.report(error));
      }
    });
    this.activeRuns.set(sessionId, { unsubscribe });
    try {
      const requestContext = new RequestContext();
      requestContext.set('caller', caller);
      requestContext.set('originalMessage', content);
      requestContext.set('freshTableRetry', freshTableRetry);
      requestContext.set('messageCreatedAt', messageCreatedAt > 0 ? messageCreatedAt : Date.now());
      const latestRelease = await this.latestReleaseJob(caller);
      if (latestRelease && content.includes(latestRelease.manifest.input.releaseId)) {
        requestContext.set('releaseJobStatus', latestRelease.status);
        requestContext.set('releaseApprovalId', latestRelease.approvalId);
      }
      await session.sendMessage({ content, requestContext, author: { id: caller.userId } });
    } catch (error) {
      await finish(false);
      throw error;
    }
  }

  private async onToolApproval(event: { toolCallId: string; toolName: string; args: unknown }, caller: Caller, session: Session) {
    if (event.toolName === 'ask_user') {
      const args = event.args as { question?: unknown; options?: Array<{ label?: unknown }> } | null;
      const question = typeof args?.question === 'string' ? args.question.trim() : '';
      const options = Array.isArray(args?.options)
        ? args.options.map(option => typeof option.label === 'string' ? option.label.trim() : '').filter(Boolean)
        : [];
      await session.respondToToolApproval({ toolCallId: event.toolCallId, decision: 'decline' });
      if (question) await this.channel.send(caller.chatId, { text: `${question}${options.length ? `\n可回复：${options.join('、')}` : ''}` });
      return;
    }
    if (officeToolNames.has(event.toolName)) {
      // Office writes create their own approval only after resolving identity and target.
      await session.respondToToolApproval({ toolCallId: event.toolCallId, decision: 'approve' });
      return;
    }
    const argumentText = JSON.stringify(event.args) ?? 'null';
    if (argumentText.length > 3_500) {
      await session.respondToToolApproval({ toolCallId: event.toolCallId, decision: 'decline' });
      throw new Error(`Approval details for ${event.toolName} exceed the card display limit`);
    }
    const keys = memoryKeys(caller);
    const proposal = await proposeAction(caller, { kind: 'mastra_tool', toolName: event.toolName, arguments: event.args }, {
      toolCallId: event.toolCallId, sessionId: keys.sessionId,
    });
    try {
      await this.channel.send(caller.chatId, { card: approvalCard(proposal.id, `${event.toolName}\n${argumentText}`) });
      const timer = setTimeout(() => {
        void expireAction(proposal.id).then(expired => {
          if (expired) return session.respondToToolApproval({ toolCallId: event.toolCallId, decision: 'decline' });
        }).catch(error => this.report(error));
      }, APPROVAL_TTL_MS + 100);
      timer.unref();
    } catch (error) {
      await (await stateDb()).execute({ sql: "UPDATE actions SET status = 'failed', result = ? WHERE id = ? AND status = 'pending'", args: [String(error), proposal.id] });
      await session.respondToToolApproval({ toolCallId: event.toolCallId, decision: 'decline' });
      throw error;
    }
  }

  async onCardAction(event: CardActionEvent) {
    const value = cardDecision.safeParse(event.action.value);
    if (!value.success) return;
    const stored = await getAction(value.data.approvalId);
    if (!stored || stored.chat_id !== event.chatId) return;
    // Older approvals have no chat_type and retain the original chat allowlist check.
    if (!isAuthorizedCaller({ userId: event.operator.openId, chatId: event.chatId, chatType: stored.chat_type ?? 'p2p' })) return;
    if (stored.kind === 'message_send' && !config.allowedUsers.has(event.operator.openId)) return;
    const decision = await decideAction(value.data.approvalId, {
      userId: event.operator.openId, chatId: event.chatId,
    }, value.data.decision === 'approve', async (action, key) => {
      if (action.kind === 'release_materials') {
        await this.channel.updateCard(event.messageId, resolvedCard('executing', `正在生成 ${action.releaseId} 的 ${action.fileCount} 个文件`));
        const run = await releaseMaterialsWorkflow.createRun();
        await (await stateDb()).execute({ sql: 'UPDATE release_jobs SET run_id = ? WHERE id = ?', args: [run.runId, action.jobId] });
        const outcome = await run.start({ inputData: { jobId: action.jobId } });
        if (outcome.status !== 'success') throw new Error(`下发材料 Workflow 结束状态：${outcome.status}`);
        return outcome.result;
      }
      const token = 'permissionMode' in action && action.permissionMode === 'user'
        ? await userAccessToken(event.operator.openId,
          action.kind.startsWith('sheet_') || action.kind.startsWith('bitable_') ? config.tableOAuthScopes : []) : undefined;
      return this.office.execute(action, key, token ?? undefined);
    }, async (stored, approved) => {
      const session = stored.session_id ? this.sessions.get(stored.session_id) : undefined;
      if (!session || !stored.tool_call_id) throw new Error('Original agent tool run has expired');
      await session.respondToToolApproval({ toolCallId: stored.tool_call_id, decision: approved ? 'approve' : 'decline' });
    });
    if (stored.kind === 'release_materials' && (decision.status === 'denied' || decision.status === 'expired')) {
      const action = officeAction.parse(JSON.parse(stored.payload));
      if (action.kind === 'release_materials') await (await stateDb()).execute({ sql: "UPDATE release_jobs SET status = ?, updated_at = ? WHERE id = ? AND status = 'pending'", args: [decision.status, Date.now(), action.jobId] });
    }
    if (decision.action && decision.status !== 'executing') {
      let detail: string | undefined;
      if (decision.status === 'succeeded' && decision.action.kind !== 'mastra_tool') {
        const action = officeAction.parse(JSON.parse(decision.action.payload));
        const result = decision.result ?? (decision.action.result ? JSON.parse(decision.action.result) : undefined);
        detail = actionResultDetail(action, result);
      }
      if (decision.status === 'failed') detail = String(decision.result ?? decision.action.result ?? '').slice(0, 1000);
      if (decision.action.kind === 'release_materials' && decision.status === 'failed') {
        const action = officeAction.parse(JSON.parse(decision.action.payload));
        if (action.kind === 'release_materials') {
          const db = await stateDb();
          const job = (await db.execute({ sql: 'SELECT status, folder_url, folder_token, error FROM release_jobs WHERE id = ?', args: [action.jobId] })).rows[0];
          const copies = (await db.execute({ sql: 'SELECT source_name, copy_url, copy_token, status FROM release_files WHERE job_id = ?', args: [action.jobId] })).rows;
          const knownCreateDenial = job?.status === 'failed' && !job.folder_token && /^Error: Feishu API (1061004|99991672):/.test(String(job.error));
          detail += `\n\n原因：${job?.error ?? '请检查 Workflow trace'}\n已生成的文件夹：${job?.folder_url ?? job?.folder_token ?? '无'}\n副本：${copies.filter(row => row.copy_token).map(row => `${row.source_name}（${row.status}）：${row.copy_url ?? row.copy_token}`).join('；') || '无'}\n${knownCreateDenial ? '创建文件夹被飞书拒绝，未产生材料。请检查应用 Drive 写入权限及材料父目录对当前群的可编辑或可管理权限，修正后重新发起。' : '请人工核对，系统不会自动重试或删除。'}`;
        }
      }
      if (detail && detail.length > 3000) {
        const full = detail;
        try {
          for (let offset = 0; offset < full.length; offset += 3000) await this.channel.send(event.chatId, { markdown: full.slice(offset, offset + 3000) });
          detail = '完整文件清单及状态已分条发送到当前会话。';
        } catch (error) {
          this.report(error);
          detail = `${full.slice(0, 2800)}\n\n详情消息发送失败，请在 Studio trace 和状态库中核对。`;
        }
      }
      const displayStatus = decision.action.kind === 'mastra_tool' && decision.status === 'succeeded' ? 'resumed' : decision.status;
      await this.channel.updateCard(event.messageId, resolvedCard(displayStatus, detail));
    }
  }

  private async runMinutesWorkflow(documentId: string, caller: Caller, replyTo: string) {
    const db = await stateDb();
    const jobId = crypto.randomUUID();
    await db.execute({ sql: 'INSERT INTO workflow_jobs(id, requester_id, chat_id, status, updated_at) VALUES (?, ?, ?, ?, ?)', args: [jobId, caller.userId, caller.chatId, 'running', Date.now()] });
    try {
      const run = await minutesToTasksWorkflow.createRun();
      await db.execute({ sql: 'UPDATE workflow_jobs SET run_id = ? WHERE id = ?', args: [run.runId, jobId] });
      const result = await run.start({ inputData: { documentId, caller } });
      if (result.status !== 'success') throw new Error(`Workflow ended with ${result.status}`);
      await db.execute({ sql: 'UPDATE workflow_jobs SET status = ?, result = ?, updated_at = ? WHERE id = ?', args: ['completed', JSON.stringify(result.result), Date.now(), jobId] });
      await this.channel.send(caller.chatId, { markdown: result.result.notificationDraft }, { replyTo });
    } catch (error) {
      await db.execute({ sql: 'UPDATE workflow_jobs SET status = ?, result = ?, updated_at = ? WHERE id = ?', args: ['failed', String(error), Date.now(), jobId] });
      throw error;
    }
  }
}
