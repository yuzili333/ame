export type SummaryRange = { startMs: number; endMs: number; label: string };

const offset = 8 * 60 * 60_000;
const datePattern = /(\d{4})\s*[年/-]\s*(\d{1,2})\s*[月/-]\s*(\d{1,2})\s*日?(?:\s*(\d{1,2}):(\d{2}))?/g;

function dateMs(parts: RegExpExecArray): number {
  const [, year, month, day, hour = '0', minute = '0'] = parts;
  const y = Number(year), m = Number(month), d = Number(day), h = Number(hour), min = Number(minute);
  const utc = Date.UTC(y, m - 1, d, h, min) - offset;
  const check = new Date(utc + offset);
  if (check.getUTCFullYear() !== y || check.getUTCMonth() + 1 !== m || check.getUTCDate() !== d || h > 23 || min > 59) throw new Error('日期或时间无效，请按“2026年10月1日 至 2026年10月15日”指定范围。');
  return utc;
}

export function parseSummaryRange(input: string, referenceMs: number): SummaryRange {
  const dates = [...input.matchAll(datePattern)];
  if (dates.length === 1 || dates.length > 2) throw new Error('请给出完整的开始和结束日期。');
  if (dates.length === 2) {
    const between = input.slice(dates[0].index! + dates[0][0].length, dates[1].index);
    if (!/^(?:\s|至|到|—|-|~|～)+$/.test(between)) throw new Error('请用“至”连接开始和结束日期。');
    const startMs = dateMs(dates[0]);
    const endMs = dateMs(dates[1]) + (dates[1][4] === undefined ? 24 * 60 * 60_000 - 1 : 59_999);
    if (startMs > endMs) throw new Error('开始时间不能晚于结束时间。');
    return { startMs, endMs, label: `${dates[0][0]} 至 ${dates[1][0]}` };
  }
  const kind = /本周|这周/.test(input) ? 'week' : /本月|这个月/.test(input) ? 'month' : /本季度|这个季度/.test(input) ? 'quarter' : /本年|今年/.test(input) ? 'year' : null;
  if (!kind) throw new Error('请指定时间范围，例如“本周”或“2026年10月1日至2026年10月15日”。');
  const now = new Date(referenceMs + offset);
  const y = now.getUTCFullYear(), m = now.getUTCMonth(), d = now.getUTCDate();
  const startMs = kind === 'week' ? Date.UTC(y, m, d - ((now.getUTCDay() + 6) % 7)) - offset
    : kind === 'month' ? Date.UTC(y, m, 1) - offset
      : kind === 'quarter' ? Date.UTC(y, Math.floor(m / 3) * 3, 1) - offset : Date.UTC(y, 0, 1) - offset;
  return { startMs, endMs: referenceMs, label: kind === 'week' ? '本周' : kind === 'month' ? '本月' : kind === 'quarter' ? '本季度' : '本年' };
}

export type SummaryRequest = { kind: 'group' | 'private'; peerName?: string; chatId?: string; range: SummaryRange };
export function parseSummaryRequest(content: string, referenceMs: number, chatType: 'p2p' | 'group'): SummaryRequest | null {
  if (!/(?:总结|汇总|梳理).{0,80}(?:聊天|群聊|私聊|单聊|对话|消息|记录)|(?:聊天|群聊|私聊|单聊|对话|消息|记录).{0,40}(?:总结|汇总|梳理)/.test(content)) return null;
  const range = parseSummaryRange(content, referenceMs);
  const nameContext = content.replace(datePattern, ' ').replace(/本周|这周|本月|这个月|本季度|这个季度|本年|今年/g, '');
  const peer = /(?:我和|我与|与)([^\s，。,；;、]{1,30}?)(?:的)?(?:私聊|单聊|聊天|对话)/.exec(nameContext);
  const privateIntent = Boolean(peer) || /私聊|单聊/.test(content) || chatType === 'p2p';
  if (privateIntent && !peer) throw new Error('请指定对话用户姓名，例如“总结本周我和张三的私聊”。');
  const chatId = /\boc_[A-Za-z0-9_-]+\b/.exec(content)?.[0];
  return { kind: privateIntent ? 'private' : 'group', peerName: peer?.[1], chatId, range };
}
