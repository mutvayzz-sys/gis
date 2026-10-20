// Runs after `wrangler deploy` in Cloudflare's build. If the build has an
// ANTHROPIC_API_KEY secret, it is copied into the Worker as a runtime secret.
// The key is passed on stdin, never printed, and never stored in the repo.
const { spawnSync } = require('child_process');
const key = (process.env.ANTHROPIC_API_KEY || '').trim();
if (!key) { console.log('put-secret: no ANTHROPIC_API_KEY in the build settings, skipping.'); process.exit(0); }
const r = spawnSync('npx', ['wrangler', 'secret', 'put', 'ANTHROPIC_API_KEY'], { input: key + '\n', stdio: ['pipe', 'inherit', 'inherit'] });
if (r.status !== 0) { console.error('put-secret: could not set the runtime secret (exit ' + r.status + ').'); process.exit(r.status || 1); }
console.log('put-secret: runtime secret ANTHROPIC_API_KEY updated.');
