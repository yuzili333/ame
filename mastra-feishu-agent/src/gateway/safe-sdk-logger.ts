import { config } from '../config';

function safeText(value: unknown): string {
  let text: string;
  if (Array.isArray(value)) text = value.map(safeText).filter(Boolean).join(' | ');
  else if (typeof value === 'string') text = value;
  else if (value instanceof Error) text = `${value.name}: ${value.message}`;
  else if (value && typeof value === 'object') {
    const item = value as Record<string, unknown>;
    const response = item.response && typeof item.response === 'object' ? item.response as Record<string, unknown> : undefined;
    const data = response?.data && typeof response.data === 'object' ? response.data as Record<string, unknown> : undefined;
    const parts = [item.message, response?.status !== undefined ? `HTTP ${response.status}` : undefined,
      item.code !== undefined ? `code ${item.code}` : undefined, item.msg,
      data?.code !== undefined ? `code ${data.code}` : undefined, data?.msg ?? data?.message];
    text = parts.filter(part => typeof part === 'string' || typeof part === 'number').map(String).join('; ')
      || `[SDK error object: ${Object.keys(item).slice(0, 8).join(', ') || 'no fields'}]`;
  } else text = String(value);
  for (const secret of [config.appSecret, config.modelApiKey, config.tokenEncryptionKey]) {
    if (secret) text = text.replaceAll(secret, '[REDACTED]');
  }
  return text.replace(/Bearer\s+[A-Za-z0-9._-]+/gi, 'Bearer [REDACTED]').slice(0, 1000);
}

export const safeSdkLogger = {
  error: (...items: unknown[]) => console.error(items.map(safeText).join(' ')),
  warn: (...items: unknown[]) => console.warn(items.map(safeText).join(' ')),
  info: (...items: unknown[]) => console.info(items.map(safeText).join(' ')),
  debug: () => undefined,
  trace: () => undefined,
};
