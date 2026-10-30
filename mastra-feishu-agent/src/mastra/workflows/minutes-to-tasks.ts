import { createStep, createWorkflow } from '@mastra/core/workflows';
import { z } from 'zod';
import { getOfficeApi } from '../../runtime/office-bridge';
import { requestOfficeApproval } from '../../runtime/approval-bridge';
import { minutesExtractor } from '../agents/agent';

const callerSchema = z.object({ userId: z.string(), chatId: z.string(), chatType: z.enum(['p2p', 'group']) });
const inputSchema = z.object({ documentId: z.string(), caller: callerSchema });
const readOutput = z.object({ documentId: z.string(), caller: callerSchema, content: z.string() });
const extraction = z.object({ summary: z.string(), tasks: z.array(z.string().min(1)).max(10) });
const extractionOutput = extraction.extend({ caller: callerSchema });
const outputSchema = z.object({ summary: z.string(), proposedTaskIds: z.array(z.string()), notificationDraft: z.string() });

const readMinutes = createStep({
  id: 'read-minutes', inputSchema, outputSchema: readOutput,
  execute: async ({ inputData }) => {
    const data = await getOfficeApi().readDocument(inputData.documentId);
    const content = typeof data.content === 'string' ? data.content : JSON.stringify(data);
    if (!content.trim()) throw new Error('The document has no readable text');
    return { ...inputData, content: content.slice(0, 50_000) };
  },
});

const extractTasks = createStep({
  id: 'extract-tasks', inputSchema: readOutput, outputSchema: extractionOutput,
  execute: async ({ inputData }) => {
    const result = await minutesExtractor.generate(`请从以下会议纪要提取简短摘要和最多 10 个明确的待办事项：\n\n${inputData.content}`, {
      structuredOutput: { schema: extraction },
    });
    const parsed = extraction.parse(result.object);
    return { ...parsed, caller: inputData.caller };
  },
});

const proposeTasks = createStep({
  id: 'propose-tasks', inputSchema: extractionOutput, outputSchema,
  execute: async ({ inputData }) => {
    const proposedTaskIds: string[] = [];
    for (const summary of inputData.tasks) {
      const proposal = await requestOfficeApproval(inputData.caller, { kind: 'task_create', summary });
      proposedTaskIds.push(proposal.approvalId);
    }
    return {
      summary: inputData.summary,
      proposedTaskIds,
      notificationDraft: `会议纪要摘要：${inputData.summary}\n待办事项：\n${inputData.tasks.map((task, index) => `${index + 1}. ${task}`).join('\n')}\n任务须逐项确认后才会创建。`,
    };
  },
});

export const minutesToTasksWorkflow = createWorkflow({
  id: 'minutes-to-tasks', inputSchema, outputSchema,
}).then(readMinutes).then(extractTasks).then(proposeTasks).commit();
