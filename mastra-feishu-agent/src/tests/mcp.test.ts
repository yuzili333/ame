import { expect, it } from 'vitest';
import { filterAllowedMcpTools } from '../mastra/mcp';

it('exposes only names explicitly allowed for a mocked MCP service', () => {
  const mockService = { internal_search: { id: 'search' }, internal_delete: { id: 'delete' }, other_search: { id: 'other' } };
  expect(filterAllowedMcpTools(mockService, { internal: { allowedTools: ['search'] } }))
    .toEqual({ internal_search: { id: 'search' } });
});
