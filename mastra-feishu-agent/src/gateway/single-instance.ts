import { mkdirSync, readFileSync, unlinkSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

function running(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try { process.kill(pid, 0); return true; }
  catch (error) { return (error as NodeJS.ErrnoException).code === 'EPERM'; }
}

export function acquireGatewayInstance(dataDir: string): () => void {
  mkdirSync(dataDir, { recursive: true });
  const lockPath = join(dataDir, 'gateway.pid');
  const owner = JSON.stringify({ pid: process.pid, startedAt: new Date().toISOString() });
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      writeFileSync(lockPath, owner, { flag: 'wx' });
      return () => {
        try {
          if (readFileSync(lockPath, 'utf8') === owner) unlinkSync(lockPath);
        } catch (error) {
          if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
        }
      };
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
      let holder: { pid?: number } = {};
      try { holder = JSON.parse(readFileSync(lockPath, 'utf8')) as { pid?: number }; }
      catch (readError) { if ((readError as NodeJS.ErrnoException).code === 'ENOENT') continue; }
      if (running(holder.pid ?? 0)) throw new Error(`飞书网关已在运行（PID ${holder.pid}）。请先停止旧进程，再启动新版本。`);
      try { unlinkSync(lockPath); }
      catch (unlinkError) { if ((unlinkError as NodeJS.ErrnoException).code !== 'ENOENT') throw unlinkError; }
    }
  }
  throw new Error('无法获取飞书网关单实例锁，请检查 data/gateway.pid');
}
