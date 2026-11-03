import { existsSync, mkdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DuckDBStore } from '@mastra/duckdb';
import { LibSQLStore } from '@mastra/libsql';

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const sourcePath = resolve(process.argv[2] ?? join(projectRoot, 'mastra.duckdb'));
const targetPath = resolve(process.argv[3] ?? join(projectRoot, 'data', 'traces.db'));

if (!existsSync(sourcePath)) {
  throw new Error(`Legacy Trace database not found: ${sourcePath}`);
}
if (sourcePath === targetPath) {
  throw new Error('The source and target Trace databases must differ');
}

mkdirSync(dirname(targetPath), { recursive: true });
const source = await new DuckDBStore({ path: sourcePath }).getStore('observability');
const target = await new LibSQLStore({ id: 'trace-migration', url: `file:${targetPath}` }).getStore('observability');
await source.init();
await target.init();

let copied = 0;
let skipped = 0;
let failed = 0;
let page = 0;
const perPage = 50;
while (true) {
  const batch = await source.listTraces({ pagination: { page, perPage } });
  for (const rootSpan of batch.spans) {
    const trace = await source.getTrace({ traceId: rootSpan.traceId });
    if (!trace) {
      failed++;
      console.error(`Trace not found during migration: ${rootSpan.traceId}`);
      continue;
    }
    for (const span of trace.spans) {
      try {
        if (await target.getSpan({ traceId: span.traceId, spanId: span.spanId })) {
          skipped++;
          continue;
        }
        await target.createSpan({ span });
        copied++;
      } catch (error) {
        failed++;
        console.error(`Failed to migrate span ${span.traceId}/${span.spanId}: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  }
  if (!batch.pagination.hasMore) break;
  page++;
}

console.log(`Trace migration complete: copied ${copied} spans, skipped ${skipped} existing spans, failed ${failed}.`);
if (failed > 0) process.exitCode = 1;
