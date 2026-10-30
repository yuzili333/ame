import type { OfficeAction } from './actions';

type CalendarCreate = Extract<OfficeAction, { kind: 'calendar_create' }>;

const shanghaiFormatter = new Intl.DateTimeFormat('en-GB', {
  timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
  hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
});

function localParts(timestampMs: number) {
  const parts = Object.fromEntries(shanghaiFormatter.formatToParts(new Date(timestampMs)).map(part => [part.type, part.value]));
  return { year: Number(parts.year), month: Number(parts.month), day: Number(parts.day), hour: parts.hour, minute: parts.minute };
}

function localDayNumber(timestampMs: number): number {
  const { year, month, day } = localParts(timestampMs);
  return Date.UTC(year, month - 1, day) / 86_400_000;
}

export function formatShanghaiTime(timestamp: string): string {
  const { year, month, day, hour, minute } = localParts(Number(timestamp) * 1000);
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')} ${hour}:${minute}`;
}

export function reconcileRelativeCalendarDate(action: CalendarCreate, message: string, receivedAt: number): CalendarCreate {
  if (!Number.isFinite(receivedAt)) return action;
  const explicitDate = /(\d{4})\s*(?:年|[-/.])\s*(\d{1,2})\s*(?:月|[-/.])\s*(\d{1,2})\s*日?/.exec(message);
  const relativeDays: Record<string, number> = { 今天: 0, 明天: 1, 后天: 2 };
  const relativeWords = [...new Set(message.match(/今天|明天|后天/g) ?? [])];
  let expectedDay: number;
  if (explicitDate) {
    const [, year, month, day] = explicitDate.map(Number);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return action;
    expectedDay = date.getTime() / 86_400_000;
  } else {
    if (relativeWords.length !== 1 || /\d{1,2}\s*月\s*\d{1,2}\s*日/.test(message)) return action;
    expectedDay = localDayNumber(receivedAt + relativeDays[relativeWords[0]] * 86_400_000);
  }
  const actualDay = localDayNumber(Number(action.startTime) * 1000);
  const dayDelta = expectedDay - actualDay;
  if (!dayDelta) return action;
  const secondsDelta = dayDelta * 86_400;
  return {
    ...action,
    startTime: String(Number(action.startTime) + secondsDelta),
    endTime: String(Number(action.endTime) + secondsDelta),
  };
}
