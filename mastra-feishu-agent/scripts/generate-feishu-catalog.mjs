// Offline catalog generation from the installed official SDK. No credentials.
import ts from 'typescript';
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
const require = createRequire(import.meta.url);
const root = dirname(require.resolve('@larksuiteoapi/node-sdk/package.json'));
const source = ts.createSourceFile('sdk.js', readFileSync(resolve(root, 'lib/index.js'), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
const types = ts.createSourceFile('sdk.d.ts', readFileSync(resolve(root, 'types/index.d.ts'), 'utf8'), ts.ScriptTarget.Latest, true);
function metadata(node, file) {
  const comment = (ts.getLeadingCommentRanges(file.text, node.pos) ?? []).map(r => file.text.slice(r.pos, r.end)).join('\n');
  const link = comment.match(/https:\/\/open\.feishu\.cn\/api-explorer\?[^\s}]+/);
  if (!link) return;
  const q = new URL(link[0]).searchParams;
  if (!q.get('project') || !q.get('resource') || !q.get('apiName')) return;
  const id = [q.get('project'), q.get('version'), q.get('resource'), q.get('apiName')].join('.');
  const description = comment.split('\n').map(l => l.replace(/^\s*\*\s?/, '').trim()).filter(l => l && !l.startsWith('/') && !l.startsWith('{@link')).join('\n').slice(0, 6000);
  return { id, description, documentation: link[0] };
}
function schema(node) {
  if (!node) return {};
  if (node.kind === ts.SyntaxKind.StringKeyword) return { type: 'string' };
  if (node.kind === ts.SyntaxKind.NumberKeyword) return { type: 'number' };
  if (node.kind === ts.SyntaxKind.BooleanKeyword) return { type: 'boolean' };
  if (ts.isArrayTypeNode(node)) return { type: 'array', items: schema(node.elementType) };
  if (ts.isParenthesizedTypeNode(node)) return schema(node.type);
  if (ts.isUnionTypeNode(node)) {
    const variants = node.types.filter(n => n.kind !== ts.SyntaxKind.UndefinedKeyword).map(schema);
    return variants.length === 1 ? variants[0] : { anyOf: variants };
  }
  if (ts.isLiteralTypeNode(node)) {
    try { return { const: JSON.parse(node.literal.getText(types)) }; } catch { return {}; }
  }
  if (ts.isTypeLiteralNode(node)) {
    const members = node.members.filter(ts.isPropertySignature);
    return { type: 'object', properties: Object.fromEntries(members.map(m => [m.name.getText(types).replace(/^['"]|['"]$/g, ''), schema(m.type)])),
      required: members.filter(m => !m.questionToken).map(m => m.name.getText(types).replace(/^['"]|['"]$/g, '')), additionalProperties: false };
  }
  return {};
}
const inputs = new Map();
function walkTypes(node) {
  if (ts.isPropertySignature(node) && node.type && ts.isFunctionTypeNode(node.type)) {
    const meta = metadata(node, types);
    const payload = node.type.parameters.find(p => p.name.getText(types) === 'payload');
    if (meta && payload) inputs.set(meta.id, schema(payload.type));
  }
  ts.forEachChild(node, walkTypes);
}
walkTypes(types);
const entries = new Map();
function walk(node) {
  if (ts.isPropertyAssignment(node)) {
    const meta = metadata(node, source);
    if (meta) {
      const body = node.initializer.getText(source);
      const path = body.match(/\$\{this\.domain\}(\/open-apis\/[^`]+)`/);
      const method = body.match(/method:\s*"(GET|POST|PATCH|PUT|DELETE)"/);
      const input = inputs.get(meta.id);
      // Files/streams and credential endpoints require purpose-built handling.
      if (path && method && input && !/responseType:\s*["'](?:stream|arraybuffer)|multipart\/form-data/.test(body) && !/\/(?:auth|authen)\//.test(path[1]) && !Object.keys(input.properties?.data?.properties ?? {}).some(key => ['file', 'file_content'].includes(key))) {
        if (input.properties) { delete input.properties.headers; input.required = (input.required ?? []).filter(x => x !== 'headers'); }
        entries.set(meta.id, { ...meta, method: method[1], path: path[1], input });
      }
    }
  }
  ts.forEachChild(node, walk);
}
walk(source);
if (entries.size < 100) throw new Error(`SDK catalog extraction failed (${entries.size} entries); inspect SDK format`);
const output = { sdkVersion: JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8')).version, operations: [...entries.values()].sort((a,b) => a.id.localeCompare(b.id)) };
writeFileSync(new URL('../src/office/feishu-api-catalog.json', import.meta.url), JSON.stringify(output));
console.log(`Generated ${entries.size} Feishu JSON API operations from SDK ${output.sdkVersion}`);
