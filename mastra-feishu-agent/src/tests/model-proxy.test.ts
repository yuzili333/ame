import { describe, expect, it } from 'vitest';
import { Agent } from '@mastra/core/agent';
import { createTool } from '@mastra/core/tools';
import { createOpenAICompatible } from '@ai-sdk/openai-compatible';
import { z } from 'zod';
import { normalizeProxyModelName } from '../model-url';

describe('OpenAI-compatible model proxy contract', () => {
  it('sends mandatory discovery to the model proxy and makes the task schema available on the next step', async () => {
    const { agent: officeAgent } = await import('../mastra/agents/agent');
    const { agentOfficeTools } = await import('../mastra/tools/office-tools');
    const { RequestContext } = await import('@mastra/core/request-context');
    const requestContext = new RequestContext();
    requestContext.set('originalMessage', '创建一个任务，主题是xxx，时间为今天18:00-19:00');
    const bodies: Array<Record<string, any>> = [];
    const provider = createOpenAICompatible({ name: 'catalog-test', baseURL: 'https://proxy.test/v1', fetch: async (_url, options) => {
      bodies.push(JSON.parse(String(options?.body)));
      const first = bodies.length === 1;
      const message = first
        ? { role: 'assistant', content: null, tool_calls: [{ id: 'discover', type: 'function', function: { name: 'office_discover_apis', arguments: '{"operationId":"office_task_create"}' } }] }
        : { role: 'assistant', content: '已检索任务创建参数。' };
      return new Response(JSON.stringify({ id: `r${bodies.length}`, object: 'chat.completion', created: 1, model: 'test', choices: [{ index: 0, message, finish_reason: first ? 'tool_calls' : 'stop' }], usage: { prompt_tokens: 1, completion_tokens: 1, total_tokens: 2 } }), { headers: { 'content-type': 'application/json' } });
    } });
    const agent = new Agent({ id: 'catalog-proxy-test', name: 'Catalog test', instructions: await officeAgent.getInstructions({ requestContext }),
      model: provider.chatModel('test'), tools: agentOfficeTools, defaultOptions: await officeAgent.getDefaultOptions({ requestContext }) });
    await agent.generate('创建一个任务，主题是xxx，时间为今天18:00-19:00');
    expect(bodies).toHaveLength(2);
    expect(bodies[0].tool_choice).toEqual({ type: 'function', function: { name: 'office_discover_apis' } });
    expect(bodies[0].tools.some((tool: any) => tool.function.name === 'office_call_api')).toBe(true);
    expect(JSON.stringify(bodies[1].messages)).toContain('office_task_create');
    expect(JSON.stringify(bodies[1].messages)).toContain('startTimestamp');
    expect(bodies[1].tool_choice).not.toEqual(bodies[0].tool_choice);
  });
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
