// A single explicit clock range is authoritative; never calculate Unix time in
// the model when the message already supplies the hours and minutes.
import { schedulingText } from './scheduling-text';
export function explicitClockRange(message: string): [number, number] | undefined {
  const text = schedulingText(message);
  const matches = [...text.matchAll(/(?<![\d:：])(上午|下午|晚上|早上|中午)?\s*(\d{1,2})\s*(?:[:：](\d{2})|点(?:(\d{1,2})分?)?)\s*(?:-|–|—|~|～|至|到)\s*(上午|下午|晚上|早上|中午)?\s*(\d{1,2})\s*(?:[:：](\d{2})|点(?:(\d{1,2})分?)?)(?!\d)/g)];
  if (matches.length !== 1) return;
  const match = matches[0];
  const minutes = (h: string, m: string | undefined, period: string | undefined) => {
    let hour = Number(h); const minute = Number(m ?? 0);
    if (hour > 23 || minute > 59) throw new Error('时间无效，请核对时分');
    if (period && /下午|晚上|中午/.test(period) && hour < 12) hour += 12;
    return hour * 60 + minute;
  };
  const start = minutes(match[2], match[3] ?? match[4], match[1]);
  const end = minutes(match[6], match[7] ?? match[8], match[5] ?? match[1]);
  if (end <= start) throw new Error('结束时间必须晚于开始时间，跨日时段请明确两端日期');
  return [start, end];
}
