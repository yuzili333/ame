import { z } from 'zod';

export const cardDecision = z.object({
  approvalId: z.string().uuid(),
  decision: z.enum(['approve', 'deny']),
});

export function approvalCard(approvalId: string, summary: string) {
  return {
    schema: '2.0',
    config: { wide_screen_mode: true },
    header: { title: { tag: 'plain_text', content: '操作确认' }, template: 'orange' },
    body: { elements: [
      { tag: 'markdown', content: `${summary}\n\n仅原请求者可确认；30 分钟后自动失效。` },
      { tag: 'column_set', flex_mode: 'none', columns: [
        { tag: 'column', width: 'weighted', weight: 1, elements: [
          { tag: 'button', text: { tag: 'plain_text', content: '确认执行' }, type: 'primary', behaviors: [{ type: 'callback', value: { approvalId, decision: 'approve' } }] },
        ] },
        { tag: 'column', width: 'weighted', weight: 1, elements: [
          { tag: 'button', text: { tag: 'plain_text', content: '拒绝' }, type: 'default', behaviors: [{ type: 'callback', value: { approvalId, decision: 'deny' } }] },
        ] },
      ] },
    ] },
  };
}

export function resolvedCard(status: string, detail?: string) {
  const labels: Record<string, string> = {
    succeeded: '已执行成功', denied: '已拒绝', expired: '确认已超时', interrupted: '服务重启，原操作已失效',
    resumed: '工具调用已放行，业务结果请查看后续消息',
    failed: '执行失败，请检查日志与目标资源后重新发起', executing: '正在执行', pending: '等待确认',
  };
  return {
    schema: '2.0',
    header: { title: { tag: 'plain_text', content: '操作状态' }, template: status === 'succeeded' ? 'green' : 'grey' },
    body: { elements: [{ tag: 'markdown', content: `${labels[status] ?? status}${detail ? `\n\n${detail}` : ''}` }] },
  };
}

export type ReleaseFileProgress = { name: string; status: string };

export function releaseProgressCard(status: string, releaseId: string, files: ReleaseFileProgress[], detail?: string) {
  const finished = status === 'succeeded' || status === 'failed' || status === 'interrupted';
  const label = (fileStatus: string) => {
    if (fileStatus === 'completed') return '更新成功';
    if (fileStatus === 'uncertain' || fileStatus === 'failed') return '更新失败（结果需核对）';
    if (finished) return fileStatus === 'pending' ? '未执行' : '更新失败（结果需核对）';
    return fileStatus === 'pending' ? '待更新' : '更新中';
  };
  const lines = files.map(file => `- ${file.name.replace(/[\r\n]/g, ' ')}：${label(file.status)}`);
  const progress = `下发编号 ${releaseId}\n文件处理进度（${files.filter(file => file.status === 'completed').length}/${files.length}）：\n${lines.join('\n')}`;
  const card = resolvedCard(status, progress);
  if (detail) card.body.elements.push({ tag: 'markdown', content: detail });
  return card;
}
