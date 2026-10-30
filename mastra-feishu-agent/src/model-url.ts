/** The OpenAI-compatible chat model appends /chat/completions itself. */
export function normalizeChatBaseURL(value: string): string {
  const url = new URL(value);
  if (url.search || url.hash) throw new Error('MODEL_PROXY_BASE_URL must not contain a query or fragment');
  let pathname = url.pathname.replace(/\/+$/, '');
  if (pathname.endsWith('/chat/completions')) {
    pathname = pathname.slice(0, -'/chat/completions'.length);
  }
  url.pathname = pathname || '/';
  return url.toString().replace(/\/$/, '');
}

/** Mastra Studio uses provider/model; the OpenAI-compatible API expects only model. */
export function normalizeProxyModelName(value: string): string {
  const name = value.trim();
  const upstream = name.startsWith('siliconflow-cn/') ? name.slice('siliconflow-cn/'.length) : name;
  if (!upstream || upstream === 'siliconflow-cn') {
    throw new Error('MODEL_PROXY_MODEL must include an upstream model name');
  }
  return upstream;
}
