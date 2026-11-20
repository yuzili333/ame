import { spawnSync } from 'node:child_process';
import * as nodeModule from 'node:module';
import { fileURLToPath } from 'node:url';
import { expect, it } from 'vitest';

// registerHooks is available in the affected Node 22.21 environment (since 22.15).
it.skipIf(!nodeModule.registerHooks)('loads the large API catalog as data under tsx, without JSON transpilation', () => {
  const result = spawnSync(process.execPath, ['--import', 'tsx', '--input-type=module', '--eval', `
    import { registerHooks } from 'node:module';
    registerHooks({ load(url, context, next) {
      const loaded = next(url, context);
      if (url.includes('feishu-api-catalog.json') && loaded.format !== 'json') {
        throw new Error('Large API catalog was passed through the TypeScript transformer');
      }
      return loaded;
    } });
    const { discoverApis, apiOperation } = await import('./src/office/openapi.ts');
    console.log(JSON.stringify({ total: discoverApis('').total,
      path: apiOperation('calendar.v4.calendar.event.attendee.create').path }));
  `], {
    cwd: fileURLToPath(new URL('../../', import.meta.url)),
    env: { ...process.env, TSX_DISABLE_CACHE: '1' }, encoding: 'utf8', timeout: 30_000,
  });
  expect(result.error).toBeUndefined();
  expect(result.status, result.stderr).toBe(0);
  const output = JSON.parse(result.stdout.trim());
  expect(output.total).toBeGreaterThan(1000);
  expect(output.path).toBe('/open-apis/calendar/v4/calendars/:calendar_id/events/:event_id/attendees');
});
