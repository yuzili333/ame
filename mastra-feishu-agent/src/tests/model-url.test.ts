import { describe, expect, it } from 'vitest';
import { normalizeChatBaseURL, normalizeProxyModelName } from '../model-url';

describe('model proxy base URL', () => {
  it('does not duplicate the chat completions route when given a full endpoint', () => {
    expect(normalizeChatBaseURL('https://api.siliconflow.cn/v1/chat/completions'))
      .toBe('https://api.siliconflow.cn/v1');
  });

  it('keeps a valid base URL unchanged', () => {
    expect(normalizeChatBaseURL('http://127.0.0.1:3001/v1/'))
      .toBe('http://127.0.0.1:3001/v1');
  });
});

describe('model names from Gateway and Studio', () => {
  it('accepts both the upstream name and Studio provider prefix', () => {
    const upstream = 'deepseek-ai/DeepSeek-V4.1-Flash';
    expect(normalizeProxyModelName(upstream)).toBe(upstream);
    expect(normalizeProxyModelName(`siliconflow-cn/${upstream}`)).toBe(upstream);
  });

  it('rejects a provider prefix without a model', () => {
    expect(() => normalizeProxyModelName('siliconflow-cn/')).toThrow('upstream model name');
  });
});
