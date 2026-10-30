import { createStep, createWorkflow } from '@mastra/core/workflows';
import { z } from 'zod';
import { ReleaseMaterialsService } from '../../office/release-materials';
import { getOfficeApi } from '../../runtime/office-bridge';

const inputSchema = z.object({ jobId: z.string().uuid() });
const outputSchema = z.object({ folderUrl: z.string().url(), files: z.array(z.object({ name: z.string(), url: z.string().url() })) });
const executeRelease = createStep({
  id: 'execute-approved-release', inputSchema, outputSchema,
  execute: async ({ inputData }) => new ReleaseMaterialsService(getOfficeApi()).executeJob(inputData.jobId),
});

export const releaseMaterialsWorkflow = createWorkflow({
  id: 'release-materials', inputSchema, outputSchema,
}).then(executeRelease).commit();
