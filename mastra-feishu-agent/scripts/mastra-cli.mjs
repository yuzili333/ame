import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

// mastra dev runs the bundled server from a different working directory.
process.env.AGENT_PROJECT_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
await import('mastra');
