import { MCPClient } from '@mastra/mcp';
import { z } from 'zod';
import { config } from '../config';

const serverSchema = z.record(z.string().regex(/^[a-z][a-z0-9_-]*$/), z.object({
  url: z.string().url(),
  allowedTools: z.array(z.string().min(1)).min(1),
}));

const serverConfig = serverSchema.parse(JSON.parse(config.mcpServersJson));
const definitions = Object.fromEntries(Object.entries(serverConfig).map(([name, value]) => {
  const url = new URL(value.url);
  if (url.protocol !== 'https:' && !(url.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(url.hostname))) {
    throw new Error(`MCP server ${name} must use HTTPS or loopback HTTP`);
  }
  return [name, { url, allowedHosts: [url.host], requireToolApproval: true, forwardInstructions: false }];
}));

export const mcpClient = Object.keys(definitions).length ? new MCPClient({ servers: definitions }) : undefined;

export function filterAllowedMcpTools<T>(tools: Record<string, T>, allowlist: Record<string, { allowedTools: string[] }>): Record<string, T> {
  const allowed = new Set(Object.entries(allowlist).flatMap(([name, value]) => value.allowedTools.map(tool => `${name}_${tool}`)));
  return Object.fromEntries(Object.entries(tools).filter(([name]) => allowed.has(name)));
}

export async function approvedMcpTools() {
  if (!mcpClient) return {};
  const discovered = await mcpClient.listTools();
  return filterAllowedMcpTools(discovered, serverConfig);
}
