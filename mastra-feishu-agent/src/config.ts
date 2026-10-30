import 'dotenv/config';
import { resolve } from 'node:path';
import { z } from 'zod';
import { normalizeChatBaseURL, normalizeProxyModelName } from './model-url';

const optionalText = z.string().trim().optional().transform(value => value || undefined);
const optionalUrl = optionalText.pipe(z.string().url().optional());
const envSchema = z.object({
  FEISHU_APP_ID: optionalText,
  FEISHU_APP_SECRET: optionalText,
  FEISHU_DOMAIN: z.string().url().optional(),
  FEISHU_DOCUMENT_HOSTS: z.string().default('feishu.cn,my.feishu.cn,larksuite.com'),
  FEISHU_TABLE_OAUTH_SCOPES: z.string().default('wiki:node:read sheets:spreadsheet:read sheets:spreadsheet:write_only bitable:app offline_access'),
  FEISHU_ALLOWED_USERS: z.string().default(''),
  FEISHU_ALLOWED_CHATS: z.string().default(''),
  FEISHU_OAUTH_AUTHORIZE_URL: optionalUrl,
  FEISHU_OAUTH_TOKEN_URL: optionalUrl,
  FEISHU_OAUTH_REDIRECT_URI: optionalUrl,
  FEISHU_OAUTH_LISTEN_HOST: z.string().default('127.0.0.1'),
  FEISHU_OAUTH_LISTEN_PORT: z.coerce.number().int().min(1).max(65535).default(4312),
  FEISHU_TOKEN_ENCRYPTION_KEY: optionalText,
  MODEL_PROXY_BASE_URL: z.string().url().default('http://127.0.0.1:3001/v1'),
  MODEL_PROXY_MODEL: z.string().min(1).default('enterprise-model'),
  MODEL_PROXY_API_KEY: optionalText,
  SILICONFLOW_CN_API_KEY: optionalText,
  AGENT_DATA_DIR: z.string().default('./data'),
  AGENT_WORKSPACE_DIR: z.string().default('./workspace'),
  MCP_SERVERS_JSON: z.string().default('{}'),
});

const env = envSchema.parse(process.env);
const csv = (value: string) => new Set(value.split(',').map(v => v.trim()).filter(Boolean));
const modelBaseURL = normalizeChatBaseURL(env.MODEL_PROXY_BASE_URL);
const usesSiliconFlowCN = new URL(modelBaseURL).hostname === 'api.siliconflow.cn';
const modelName = normalizeProxyModelName(env.MODEL_PROXY_MODEL);
if (usesSiliconFlowCN && env.MODEL_PROXY_API_KEY && env.SILICONFLOW_CN_API_KEY
  && env.MODEL_PROXY_API_KEY !== env.SILICONFLOW_CN_API_KEY) {
  throw new Error('MODEL_PROXY_API_KEY and SILICONFLOW_CN_API_KEY must match for the SiliconFlow CN model used by Gateway and Studio');
}

export const config = {
  appId: env.FEISHU_APP_ID,
  appSecret: env.FEISHU_APP_SECRET,
  domain: env.FEISHU_DOMAIN,
  documentHosts: [...new Set([...env.FEISHU_DOCUMENT_HOSTS.split(',').map(v => v.trim().toLowerCase()).filter(Boolean), ...(env.FEISHU_DOMAIN ? [new URL(env.FEISHU_DOMAIN).hostname.toLowerCase()] : [])])],
  tableOAuthScopes: env.FEISHU_TABLE_OAUTH_SCOPES.split(/\s+/).filter(Boolean),
  allowedUsers: csv(env.FEISHU_ALLOWED_USERS),
  allowedChats: csv(env.FEISHU_ALLOWED_CHATS),
  oauthAuthorizeUrl: env.FEISHU_OAUTH_AUTHORIZE_URL,
  oauthTokenUrl: env.FEISHU_OAUTH_TOKEN_URL,
  oauthRedirectUri: env.FEISHU_OAUTH_REDIRECT_URI,
  oauthListenHost: env.FEISHU_OAUTH_LISTEN_HOST,
  oauthListenPort: env.FEISHU_OAUTH_LISTEN_PORT,
  tokenEncryptionKey: env.FEISHU_TOKEN_ENCRYPTION_KEY,
  modelBaseURL,
  modelName,
  studioModelName: usesSiliconFlowCN || env.MODEL_PROXY_MODEL.trim().startsWith('siliconflow-cn/')
    ? `siliconflow-cn/${modelName}` : modelName,
  modelApiKey: env.MODEL_PROXY_API_KEY ?? (usesSiliconFlowCN ? env.SILICONFLOW_CN_API_KEY : undefined),
  dataDir: resolve(env.AGENT_DATA_DIR),
  workspaceDir: resolve(env.AGENT_WORKSPACE_DIR),
  mcpServersJson: env.MCP_SERVERS_JSON,
};

export function requireFeishuConfig() {
  if (!config.appId || !config.appSecret || !config.domain) {
    throw new Error('FEISHU_APP_ID, FEISHU_APP_SECRET and FEISHU_DOMAIN are required for the gateway');
  }
  return { appId: config.appId, appSecret: config.appSecret, domain: config.domain };
}
