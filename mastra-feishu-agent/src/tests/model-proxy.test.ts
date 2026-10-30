import { describe, expect, it } from 'vitest';
import { Agent } from '@mastra/core/agent';
import { createTool } from '@mastra/core/tools';
import { createOpenAICompatible } from '@ai-sdk/openai-compatible';
import { z } from 'zod';
import { normalizeProxyModelName } from '../model-url';

describe('OpenAI-compatible model proxy contract', () => {
  it('accepts a mocked tool_calls response and executes a typed tool', async () => {
    let calls = 0;
    let toolsCalled = 0;
    const sentModels: string[] = [];
    const provider = createOpenAICompatible({
      name: 'mock-proxy', baseURL: 'http://127.0.0.1:3001/v1',
      fetch: async (_url, options) => {
        sentModels.push((JSON.parse(String(options?.body)) as { model: string }).model);
        calls++;
        const message = calls === 1
          ? { role: 'assistant', content: null, tool_calls: [{ id: 'call_1', type: 'function', function: { name: 'lookup', arguments: '{"query":"hello"}' } }] }
          : { role: 'assistant', content: 'Found the item.' };
        return new Response(JSON.stringify({ id: `mock-${calls}`, object: 'chat.completion', created: 1, model: 'mock-model', choices: [{ index: 0, message, finish_reason: calls === 1 ? 'tool_calls' : 'stop' }], usage: { prompt_tokens: 1, completion_tokens: 1, total_tokens: 2 } }), { status: 200, headers: { 'content-type': 'application/json' } });
      },
    });
    const agent = new Agent({
      id: 'proxy-test-agent', name: 'Proxy test', instructions: 'Use lookup.',
      model: provider.chatModel(normalizeProxyModelName('siliconflow-cn/mock-model')),
      tools: { lookup: createTool({ id: 'lookup', description: 'Look up a word', inputSchema: z.object({ query: z.string() }), execute: async ({ query }) => { toolsCalled++; return { query }; } }) },
    });
    const result = await agent.generate('Find hello', { maxSteps: 2 });
    expect(result.text).toContain('Found the item');
    expect(calls).toBe(2);
    expect(toolsCalled).toBe(1);
    expect(sentModels).toEqual(['mock-model', 'mock-model']);
  });
});
