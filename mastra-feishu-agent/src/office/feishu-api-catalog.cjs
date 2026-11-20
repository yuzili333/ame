// Native JSON loading avoids tsx transforming this >1 MiB catalog through
// esbuild's temporary-file handoff, which can fail with Access denied on Windows.
// Keep require static so Mastra can still include the data in deployment bundles.
module.exports = require('./feishu-api-catalog.json');
