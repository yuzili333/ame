import { expect, it, vi } from 'vitest';
import { safeSdkLogger } from '../gateway/safe-sdk-logger';

it('shows useful SDK error details from a formatted error array without logging request credentials', () => {
  const logged = vi.spyOn(console, 'error').mockImplementation(() => undefined);
  try {
    safeSdkLogger.error([
      { message: 'Request failed with status code 403', response: { status: 403, data: { code: 99991672, msg: 'Access denied' } }, config: { data: { app_secret: 'never-print-me' } } },
      { code: 99991672, msg: 'Access denied' },
    ]);
    const output = String(logged.mock.calls[0]?.[0]);
    expect(output).toContain('403');
    expect(output).toContain('99991672');
    expect(output).toContain('Access denied');
    expect(output).not.toContain('never-print-me');
    expect(output).not.toContain('[SDK error object]');
  } finally {
    logged.mockRestore();
  }
});
