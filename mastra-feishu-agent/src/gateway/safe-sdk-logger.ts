import { config } from '../config';

function safeText(value: unknown): string {
  let text: string;
  if (typeof value === 'string') text = value;
  else if (value instanceof Error) text = `${value.name}: ${value.message}`;
  else if (value && typeof value === 'object') {
    const item = value as { code?: unknown; msg?: unknown; message?: unknown };
    text = [item.code, item.msg ?? item.message].filter(part => part !== undefined).map(String).join(': ') || '[SDK error object]';
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
