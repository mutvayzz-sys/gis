// Serves the page and relays the assistant's requests to Claude.
// The API key lives only in the Cloudflare secret ANTHROPIC_API_KEY. Nothing is stored.
import PROMPT from '../src/prompt.txt';

const TOOLS = new Set(['get_overview', 'read_cells', 'choose_sheet', 'set_header_row', 'set_column', 'set_rules', 'edit_number', 'set_cell', 'choose_option', 'go_to_step', 'set_cells', 'enter_rows', 'set_phase', 'add_rule', 'remove_rule', 'set_simulation_limits', 'set_simulated_numbers']);
const MAX_BODY = 400 * 1024;

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json', 'cache-control': 'no-store' } });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/api/health') return json({ ok: Boolean(env.ANTHROPIC_API_KEY) });

    if (url.pathname === '/api/chat') {
      if (request.method !== 'POST') return json({ error: 'Use POST.' }, 405);
      const origin = request.headers.get('Origin');
      if (origin && new URL(origin).host !== url.host) return json({ error: 'Requests are only accepted from this site.' }, 403);
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
