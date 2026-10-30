import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { acquireGatewayInstance } from '../gateway/single-instance';

it('rejects a second live gateway and releases its PID lock', () => {
  const dir = mkdtempSync(join(tmpdir(), 'feishu-gateway-lock-'));
  const release = acquireGatewayInstance(dir);
  expect(JSON.parse(readFileSync(join(dir, 'gateway.pid'), 'utf8')).pid).toBe(process.pid);
  expect(() => acquireGatewayInstance(dir)).toThrow(`PID ${process.pid}`);
  release();
  const second = acquireGatewayInstance(dir);
  second();
});

it('replaces a stale gateway PID lock', () => {
  const dir = mkdtempSync(join(tmpdir(), 'feishu-gateway-stale-lock-'));
  writeFileSync(join(dir, 'gateway.pid'), JSON.stringify({ pid: 0, startedAt: 'old' }));
  const release = acquireGatewayInstance(dir);
  expect(JSON.parse(readFileSync(join(dir, 'gateway.pid'), 'utf8')).pid).toBe(process.pid);
  release();
});
