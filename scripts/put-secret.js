// Runs after `wrangler deploy` in Cloudflare's build. Any of these found in the
// build's secrets are copied into the Worker as runtime secrets. Values go in on
// stdin, are never printed, and are never stored in the repo.
const { spawnSync } = require('child_process');
const NAMES = ['ANTHROPIC_API_KEY', 'GIS_GITHUB_TOKEN'];
let failed = 0;
for (const name of NAMES) {
  const value = (process.env[name] || '').trim();
  if (!value) { console.log('put-secret: ' + name + ' not in the build settings, skipping.'); continue; }
  const r = spawnSync('npx', ['wrangler', 'secret', 'put', name], { input: value + '\n', stdio: ['pipe', 'inherit', 'inherit'] });
  if (r.status !== 0) { console.error('put-secret: could not set ' + name + ' (exit ' + r.status + ').'); failed = r.status || 1; }
  else console.log('put-secret: runtime secret ' + name + ' updated.');
}
process.exit(failed);
