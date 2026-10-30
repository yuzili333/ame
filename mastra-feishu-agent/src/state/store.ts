import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { createClient, type Client } from '@libsql/client';
import { config } from '../config';

export type ActionStatus = 'pending' | 'executing' | 'succeeded' | 'failed' | 'denied' | 'expired' | 'interrupted';
export type StoredAction = {
  id: string;
  kind: string;
  payload: string;
  requester_id: string;
  chat_id: string;
  chat_type: 'p2p' | 'group' | null;
  status: ActionStatus;
  created_at: number;
  expires_at: number;
  result: string | null;
  tool_call_id: string | null;
  session_id: string | null;
  origin_message_id: string | null;
  card_message_id: string | null;
};

let clientPromise: Promise<Client> | undefined;

export async function stateDb(): Promise<Client> {
  if (clientPromise) return clientPromise;
  clientPromise = initializeStateDb().catch(error => {
    clientPromise = undefined;
    throw error;
  });
  return clientPromise;
}

async function initializeStateDb(): Promise<Client> {
  mkdirSync(config.dataDir, { recursive: true });
  const db = createClient({ url: `file:${join(config.dataDir, 'state.db')}` });
  await db.execute(`CREATE TABLE IF NOT EXISTS actions (
    id TEXT PRIMARY KEY,
    kind TEXT NOT NULL,
    payload TEXT NOT NULL,
    requester_id TEXT NOT NULL,
    chat_id TEXT NOT NULL,
    chat_type TEXT,
    status TEXT NOT NULL,
    created_at INTEGER NOT NULL,
    expires_at INTEGER NOT NULL,
    result TEXT,
    tool_call_id TEXT,
    session_id TEXT
  )`);
  const actionColumns = await db.execute('PRAGMA table_info(actions)');
  if (!actionColumns.rows.some(row => row.name === 'chat_type')) {
    try {
      await db.execute('ALTER TABLE actions ADD COLUMN chat_type TEXT');
    } catch (error) {
      if (!(error instanceof Error) || !error.message.includes('duplicate column name')) throw error;
    }
  }
  if (!actionColumns.rows.some(row => row.name === 'origin_message_id')) await db.execute('ALTER TABLE actions ADD COLUMN origin_message_id TEXT');
  if (!actionColumns.rows.some(row => row.name === 'card_message_id')) await db.execute('ALTER TABLE actions ADD COLUMN card_message_id TEXT');
  await db.execute(`CREATE TABLE IF NOT EXISTS seen_messages (
    message_id TEXT PRIMARY KEY,
    seen_at INTEGER NOT NULL
  )`);
  await db.execute(`CREATE TABLE IF NOT EXISTS recalled_messages (
    message_id TEXT PRIMARY KEY, chat_id TEXT NOT NULL, recalled_at INTEGER NOT NULL
  )`);
  await db.execute(`CREATE TABLE IF NOT EXISTS queued_messages (
    message_id TEXT PRIMARY KEY, chat_id TEXT NOT NULL, message_json TEXT NOT NULL,
    status TEXT NOT NULL, created_at INTEGER NOT NULL
  )`);
  await db.execute('CREATE INDEX IF NOT EXISTS queued_messages_chat ON queued_messages(chat_id, status, created_at)');
  await db.execute(`CREATE TABLE IF NOT EXISTS workflow_jobs (
    id TEXT PRIMARY KEY,
    run_id TEXT,
    requester_id TEXT NOT NULL,
    chat_id TEXT NOT NULL,
    status TEXT NOT NULL,
    result TEXT,
    updated_at INTEGER NOT NULL
  )`);
  await db.execute(`CREATE TABLE IF NOT EXISTS release_jobs (
    id TEXT PRIMARY KEY, requester_id TEXT NOT NULL, chat_id TEXT NOT NULL,
    status TEXT NOT NULL, manifest_json TEXT NOT NULL, run_id TEXT,
    folder_token TEXT, folder_url TEXT, result_json TEXT, error TEXT, approval_id TEXT, request_key TEXT,
    created_at INTEGER NOT NULL, updated_at INTEGER NOT NULL
  )`);
  const releaseColumns = await db.execute('PRAGMA table_info(release_jobs)');
  if (!releaseColumns.rows.some(row => row.name === 'origin_message_id')) await db.execute('ALTER TABLE release_jobs ADD COLUMN origin_message_id TEXT');
  if (!releaseColumns.rows.some(row => row.name === 'approval_id')) await db.execute('ALTER TABLE release_jobs ADD COLUMN approval_id TEXT');
  if (!releaseColumns.rows.some(row => row.name === 'request_key')) await db.execute('ALTER TABLE release_jobs ADD COLUMN request_key TEXT');
  await db.execute('CREATE UNIQUE INDEX IF NOT EXISTS release_jobs_request_key ON release_jobs(request_key)');
  await db.execute(`CREATE TABLE IF NOT EXISTS release_files (
    job_id TEXT NOT NULL, source_token TEXT NOT NULL, source_name TEXT NOT NULL,
    type TEXT NOT NULL, status TEXT NOT NULL, copy_token TEXT, copy_url TEXT,
    error TEXT, PRIMARY KEY(job_id, source_token)
  )`);
  await db.execute(`CREATE TABLE IF NOT EXISTS permission_requests (
    id TEXT PRIMARY KEY,
    requester_id TEXT NOT NULL,
    chat_id TEXT NOT NULL,
    payload TEXT NOT NULL,
    status TEXT NOT NULL,
    created_at INTEGER NOT NULL,
    expires_at INTEGER NOT NULL,
    chat_type TEXT
  )`);
  const permissionColumns = await db.execute('PRAGMA table_info(permission_requests)');
  if (!permissionColumns.rows.some(row => row.name === 'chat_type')) await db.execute('ALTER TABLE permission_requests ADD COLUMN chat_type TEXT');
  if (!permissionColumns.rows.some(row => row.name === 'origin_message_id')) await db.execute('ALTER TABLE permission_requests ADD COLUMN origin_message_id TEXT');
  await db.execute(`CREATE INDEX IF NOT EXISTS permission_requests_lookup ON permission_requests(requester_id, chat_id, status)`);
  await db.execute(`CREATE TABLE IF NOT EXISTS user_tokens (
    user_id TEXT PRIMARY KEY,
    encrypted_token TEXT NOT NULL,
    encrypted_refresh_token TEXT,
    expires_at INTEGER NOT NULL,
    updated_at INTEGER NOT NULL,
    scopes TEXT
  )`);
  const tokenColumns = await db.execute('PRAGMA table_info(user_tokens)');
  if (!tokenColumns.rows.some(row => row.name === 'scopes')) await db.execute('ALTER TABLE user_tokens ADD COLUMN scopes TEXT');
  return db;
}

export async function isMessageRecalled(messageId?: string): Promise<boolean> {
  if (!messageId) return false;
  const rows = await (await stateDb()).execute({ sql: 'SELECT 1 FROM recalled_messages WHERE message_id = ?', args: [messageId] });
  return rows.rows.length > 0;
}

export async function markMessageSeen(messageId: string): Promise<boolean> {
  const db = await stateDb();
  const result = await db.execute({ sql: 'INSERT OR IGNORE INTO seen_messages(message_id, seen_at) VALUES (?, ?)', args: [messageId, Date.now()] });
  return result.rowsAffected === 1;
}

export async function recoverInterruptedApprovals(): Promise<number> {
  const db = await stateDb();
  const result = await db.execute({
    sql: "UPDATE actions SET status = 'interrupted' WHERE status = 'executing' OR (status = 'pending' AND kind = 'mastra_tool')",
    args: [],
  });
  return result.rowsAffected;
}

export async function recoverInterruptedWorkflows(): Promise<number> {
  const db = await stateDb();
  const result = await db.execute({
    sql: "UPDATE workflow_jobs SET status = 'interrupted', updated_at = ? WHERE status = 'running'",
    args: [Date.now()],
  });
  await db.execute({ sql: "UPDATE release_jobs SET status = 'needs_review', error = '服务重启：上次执行结果可能不明，请人工核对，不会自动重试', updated_at = ? WHERE status IN ('running', 'creating_folder')", args: [Date.now()] });
  return result.rowsAffected;
}

export async function getAction(id: string): Promise<StoredAction | null> {
  const db = await stateDb();
  const result = await db.execute({ sql: 'SELECT * FROM actions WHERE id = ?', args: [id] });
  return (result.rows[0] as unknown as StoredAction | undefined) ?? null;
}
