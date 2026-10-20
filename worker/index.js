// Serves the page and relays the assistant's requests to Claude.
// Secrets live only in Cloudflare: ANTHROPIC_API_KEY, and GIS_GITHUB_TOKEN for reports.
// Nothing about a sheet is stored. /api/report opens a GitHub issue when the assistant
// hits something it cannot do; it carries a short summary only, never sheet data.
import PROMPT from '../src/prompt.txt';

const TOOLS = new Set(['get_overview', 'read_cells', 'choose_sheet', 'set_header_row', 'set_column', 'set_rules', 'edit_number', 'set_cell', 'choose_option', 'go_to_step', 'set_cells', 'enter_rows', 'set_phase', 'add_rule', 'remove_rule', 'set_simulation_limits', 'set_simulated_numbers', 'report_gap']);
const MAX_BODY = 400 * 1024;
const REPORT_REPO = 'mutvayzz-sys/gis';

async function allowed(limiter, key) {
  if (!limiter || typeof limiter.limit !== 'function') return true;
  try { const { success } = await limiter.limit({ key }); return success; } catch { return true; }
}
function clean(v, n) {
  // Plain text only; stop @mentions and very long input.
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]+/g, ' ').replace(/@/g, '@\u200b').trim().slice(0, n);
}
function sameSite(request, url) {
  const origin = request.headers.get('Origin');
  return !origin || new URL(origin).host === url.host;
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json', 'cache-control': 'no-store' } });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/api/health') return json({ ok: Boolean(env.ANTHROPIC_API_KEY), reports: Boolean(env.GIS_GITHUB_TOKEN) });

    if (url.pathname === '/api/report') {
      if (request.method !== 'POST') return json({ error: 'Use POST.' }, 405);
      if (!sameSite(request, url)) return json({ error: 'Requests are only accepted from this site.' }, 403);
      if (!env.GIS_GITHUB_TOKEN) return json({ error: 'Reports are not set up.' }, 503);
      if (Number(request.headers.get('content-length') || 0) > 8 * 1024) return json({ error: 'That report is too large.' }, 413);
      const ip = request.headers.get('cf-connecting-ip') || 'unknown';
      if (!(await allowed(env.REPORT_LIMIT, ip)) || !(await allowed(env.REPORT_LIMIT, 'all'))) return json({ error: 'Too many reports right now.' }, 429);
      let b;
      try { b = await request.json(); } catch { return json({ error: 'The report could not be read.' }, 400); }
      const asked = clean(b.asked, 400), why = clean(b.why, 600), step = clean(b.step, 40), kind = clean(b.kind, 40);
      if (!asked) return json({ error: 'Nothing to report.' }, 400);
      const body = [
        '**What the person asked for**', asked, '',
        '**Why the assistant could not do it**', why || '(not given)', '',
        '| | |', '|---|---|', '| Kind | ' + (kind || 'missing feature') + ' |', '| Page step | ' + (step || '?') + ' |', '| When | ' + new Date().toISOString() + ' |', '',
        '_Sent automatically by the assistant on gis.gcaplabs.com. It is written to hold no student data, but check before sharing this issue._'
      ].join('\n');
      const res = await fetch('https://api.github.com/repos/' + (env.REPORT_REPO || REPORT_REPO) + '/issues', {
        method: 'POST',
        headers: { authorization: 'Bearer ' + env.GIS_GITHUB_TOKEN, accept: 'application/vnd.github+json', 'x-github-api-version': '2022-11-28', 'user-agent': 'gis-assistant', 'content-type': 'application/json' },
        body: JSON.stringify({ title: 'Assistant could not: ' + asked.slice(0, 90), body, labels: ['assistant-gap'] })
      });
      if (!res.ok) return json({ error: 'The report could not be sent (' + res.status + ').' }, 502);
      const issue = await res.json().catch(() => ({}));
      return json({ ok: true, number: issue.number || null });
    }

    if (url.pathname === '/api/chat') {
      if (request.method !== 'POST') return json({ error: 'Use POST.' }, 405);
      if (!sameSite(request, url)) return json({ error: 'Requests are only accepted from this site.' }, 403);
      if (!(await allowed(env.CHAT_LIMIT, request.headers.get('cf-connecting-ip') || 'unknown'))) return json({ error: 'Too many requests right now. Wait a minute and try again.' }, 429);
      if (!env.ANTHROPIC_API_KEY) return json({ error: 'The assistant is not set up yet.' }, 503);
      if (Number(request.headers.get('content-length') || 0) > MAX_BODY) return json({ error: 'That request is too large.' }, 413);

      let body;
      try { body = await request.json(); } catch { return json({ error: 'The request could not be read.' }, 400); }
      const messages = Array.isArray(body.messages) ? body.messages.slice(-40) : null;
      if (!messages || !messages.length) return json({ error: 'No message was sent.' }, 400);
      const tools = (Array.isArray(body.tools) ? body.tools : [])
        .filter((t) => t && TOOLS.has(t.name))
        .map((t) => ({ name: t.name, description: String(t.description || '').slice(0, 1024), input_schema: t.input_schema || { type: 'object', properties: {} } }));

      const res = await fetch(env.ANTHROPIC_URL || 'https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'x-api-key': env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
        body: JSON.stringify({ model: env.MODEL || 'claude-sonnet-5-5', max_tokens: 2048, system: PROMPT, messages, tools })
      });
      let data = {};
      try { data = await res.json(); } catch { /* fall through */ }
      if (!res.ok) return json({ error: (data.error && data.error.message) || 'Claude could not answer right now. Try again.' }, 502);
      return json({ content: data.content || [], stop_reason: data.stop_reason });
    }

    return env.ASSETS.fetch(request);
  }
};
