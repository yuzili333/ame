import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { Mastra } from '@mastra/core/mastra';
import { AgentController } from '@mastra/core/agent-controller';
import { LibSQLStore } from '@mastra/libsql';
import { DuckDBStore } from '@mastra/duckdb';
import { MastraCompositeStore } from '@mastra/core/storage';
import { MastraStorageExporter, Observability, SensitiveDataFilter } from '@mastra/observability';
import { config } from '../config';
import { agent, minutesExtractor, workspace } from './agents/agent';
import { minutesToTasksWorkflow } from './workflows/minutes-to-tasks';
import { releaseMaterialsWorkflow } from './workflows/release-materials';
import { officeTools } from './tools/office-tools';
import { startScheduleTool, stopScheduleTool } from './tools/schedule-tools';

mkdirSync(config.dataDir, { recursive: true });
mkdirSync(config.workspaceDir, { recursive: true });
const storageUrl = `file:${join(config.dataDir, 'mastra.db')}`;

export const controller = new AgentController({
  id: 'office-controller',
  agent,
  storage: new LibSQLStore({ id: 'controller-storage', url: storageUrl }),
  workspace,
  modes: [{ id: 'office', name: 'Office', metadata: { default: true }, instructions: 'Assist with approved Feishu office work.' }],
  toolCategoryResolver: toolName => {
    if (toolName.includes('write_file') || toolName.includes('edit_file') || toolName.includes('delete') || toolName.includes('execute_command') || toolName.includes('schedule')) return 'execute';
    return null;
  },
});

export const mastra = new Mastra({
  server: { host: '127.0.0.1', studioHost: '127.0.0.1' },
  bundler: { externals: ['@duckdb/node-bindings'] },
  agents: { agent, minutesExtractor },
  agentControllers: { controller },
  workflows: { minutesToTasksWorkflow, releaseMaterialsWorkflow },
  tools: { ...officeTools, startScheduleTool, stopScheduleTool },
  storage: new MastraCompositeStore({
    id: 'composite-storage',
    default: new LibSQLStore({ id: 'mastra-storage', url: storageUrl }),
    domains: { observability: await new DuckDBStore().getStore('observability') },
  }),
  observability: new Observability({
    configs: { default: { serviceName: 'feishu-office-agent', exporters: [new MastraStorageExporter()], spanOutputProcessors: [new SensitiveDataFilter()] } },
  }),
});
