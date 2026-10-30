import { config } from '../config';
import { acquireGatewayInstance } from './single-instance';

const release = acquireGatewayInstance(config.dataDir);
process.once('exit', release);
const { FeishuGateway } = await import('./gateway');
const gateway = new FeishuGateway();
try { await gateway.start(); }
catch (error) { release(); throw error; }
console.log(`Feishu gateway connected. PID ${process.pid}; entry ${import.meta.url}; Mastra Studio: http://127.0.0.1:4111`);

for (const signal of ['SIGINT', 'SIGTERM'] as const) {
  process.once(signal, () => { void gateway.stop().finally(() => { release(); process.exit(0); }); });
}
