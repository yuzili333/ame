import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

// Keep local credentials out of these configuration tests.
vi.mock('dotenv/config', () => ({}));

beforeEach(() => {
  vi.resetModules();
});

afterEach(() => {
  vi.unstubAllEnvs();
});

describe.each(['http://model-proxy.internal/v1', 'https://api.siliconflow.cn/v1'])('model credentials for %s', baseURL => {
  it.each([
    { name: 'prefers the proxy key when both keys differ', proxy: 'proxy-key', siliconflow: 'siliconflow-key', expected: 'proxy-key' },
    { name: 'accepts only a proxy key', proxy: 'proxy-key', siliconflow: undefined, expected: 'proxy-key' },
    { name: 'falls back when the proxy key is absent', proxy: undefined, siliconflow: 'siliconflow-key', expected: 'siliconflow-key' },
    { name: 'falls back when the proxy key is empty', proxy: '', siliconflow: 'siliconflow-key', expected: 'siliconflow-key' },
    { name: 'falls back when the proxy key is whitespace', proxy: '  ', siliconflow: ' siliconflow-key ', expected: 'siliconflow-key' },
    { name: 'trims the preferred proxy key', proxy: ' proxy-key ', siliconflow: 'siliconflow-key', expected: 'proxy-key' },
    { name: 'allows a proxy without API key authentication', proxy: undefined, siliconflow: undefined, expected: undefined },
    { name: 'ignores empty credentials', proxy: '', siliconflow: '  ', expected: undefined },
  ])('$name', async ({ proxy, siliconflow, expected }) => {
    vi.stubEnv('MODEL_PROXY_BASE_URL', baseURL);
    vi.stubEnv('MODEL_PROXY_MODEL', 'enterprise-model');
    vi.stubEnv('MODEL_PROXY_API_KEY', proxy);
    vi.stubEnv('SILICONFLOW_CN_API_KEY', siliconflow);

    const { config } = await import('../config');
    expect(config.modelApiKey).toBe(expected);
    expect(config.modelBaseURL).toBe(baseURL);
  });
});
