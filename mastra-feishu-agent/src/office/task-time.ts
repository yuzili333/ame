import type { OfficeAction } from './actions';
import { formatShanghaiTime } from './calendar-time';

type TaskWrite = Extract<OfficeAction, { kind: 'task_create' | 'task_update' }>;
const dayMs = 86_400_000;
const offsetMs = 8 * 3_600_000;
const localDay = (ms: number) => Math.floor((ms + offsetMs) / dayMs);

export function assertTaskTimeRange(action: { startTimestamp?: string; dueTimestamp?: string }) {
  if (action.startTimestamp && action.dueTimestamp && Number(action.startTimestamp) > Number(action.dueTimestamp)) {
    throw new Error('任务截止时间不能早于开始时间，请明确开始和截止日期时间');
  }
}

// Resolve only a single unambiguous date/range. The original message timestamp
// is trusted gateway context and must survive OAuth, including across midnight.
export function reconcileTaskTime<T extends TaskWrite>(action: T, message: string, receivedAt: number): T {
  if (!Number.isFinite(receivedAt)) return action;
  const scheduling = message.split(/[，,。；;\n]\s*(?:主题|标题|名称|说明|描述)\s*(?:是|为|[:：])/)[0];
  const dates = [...scheduling.matchAll(/(\d{4})\s*(?:年|[-/.])\s*(\d{1,2})\s*(?:月|[-/.])\s*(\d{1,2})\s*日?/g)];
  const relative = [...new Set(scheduling.match(/今天|明天|后天/g) ?? [])];
  let expectedDay: number;
  if (dates.length === 1 && relative.length === 0) {
    const [, y, m, d] = dates[0].map(Number);
    const date = new Date(Date.UTC(y, m - 1, d));
    if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) throw new Error('任务日期无效，请核对年月日');
    expectedDay = date.getTime() / dayMs;
  } else if (dates.length === 0 && relative.length === 1 && !/(?<![\d:：])\d{1,2}\s*(?:月|[-/.])\s*\d{1,2}(?![\d:：])/.test(scheduling)) {
    expectedDay = localDay(receivedAt) + ({ 今天: 0, 明天: 1, 后天: 2 }[relative[0]] ?? 0);
  } else return action;

  const result = { ...action };
  const ranges = [...scheduling.matchAll(/(?<![\d:：])(\d{1,2})[:：](\d{2})\s*(?:-|–|—|~|～|至|到)\s*(\d{1,2})[:：](\d{2})(?!\d)/g)];
  if (ranges.length > 1) return action;
  if (ranges.length === 1) {
    const [, h1, m1, h2, m2] = ranges[0].map(Number);
    if (h1 > 23 || h2 > 23 || m1 > 59 || m2 > 59) throw new Error('任务时间无效，请使用有效的时分');
    // Respect an explicit afternoon marker for a 12-hour range.
    const afternoon = /下午|晚上|傍晚/.test(scheduling.slice(0, ranges[0].index));
    const hour = (h: number) => afternoon && h < 12 ? h + 12 : h;
    result.startTimestamp = String(expectedDay * dayMs - offsetMs + (hour(h1) * 60 + m1) * 60_000);
    result.dueTimestamp = String(expectedDay * dayMs - offsetMs + (hour(h2) * 60 + m2) * 60_000);
  } else {
    const anchor = action.startTimestamp ?? action.dueTimestamp;
    if (!anchor) return action;
    const delta = (expectedDay - localDay(Number(anchor))) * dayMs;
    if (action.startTimestamp) result.startTimestamp = String(Number(action.startTimestamp) + delta);
    if (action.dueTimestamp) result.dueTimestamp = String(Number(action.dueTimestamp) + delta);
  }
  assertTaskTimeRange(result);
  // Correct only a generated scheduling line, never dates supplied by the user
  // in task content. Structural start/due fields carry the actual schedule.
  if (result.description && result.startTimestamp && result.dueTimestamp) {
    result.description = result.description.replace(/^时间窗口[：:][ \t]*\d{4}-\d{2}-\d{2}[ \t]+\d{1,2}:\d{2}[ \t]*[-–—~～至到][ \t]*(?:\d{4}-\d{2}-\d{2}[ \t]+)?\d{1,2}:\d{2}(?:[ \t]*[（(]Asia\/Shanghai[）)])?[ \t]*$/gm, line => message.includes(line) ? line
      : `时间窗口：${formatShanghaiTime(String(Number(result.startTimestamp) / 1000))}—${formatShanghaiTime(String(Number(result.dueTimestamp) / 1000))}（Asia/Shanghai）`);
  }
  return result;
}
