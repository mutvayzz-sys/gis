// Serves the page and relays the assistant's requests to Claude.
// Secrets live only in Cloudflare: ANTHROPIC_API_KEY, and GIS_GITHUB_TOKEN for reports.
// Nothing about a sheet is stored here. For the assistant, the spreadsheet is put in Claude's
// Python sandbox (Anthropic keeps sandboxes up to 30 days); the uploaded copy is deleted at once.
// /api/report opens a GitHub issue when the assistant hits something it cannot do; it carries a
// short summary only, never sheet data.
import PROMPT from '../src/prompt.txt';
import puppeteer from '@cloudflare/puppeteer';

const TOOLS = new Set(['get_overview', 'read_cells', 'choose_sheet', 'set_header_row', 'set_column', 'set_rules', 'edit_number', 'set_cell', 'choose_option', 'go_to_step', 'set_cells', 'enter_rows', 'set_phase', 'add_rule', 'remove_rule', 'set_simulation_limits', 'set_simulated_numbers', 'report_gap', 'undo_last_change', 'offer_downloads', 'use_all_sheets', 'scan_sheet']);
const MAX_BODY = 14 * 1024 * 1024; // room for a few attached pictures and the spreadsheet itself
const API = 'https://api.anthropic.com/v1';
const SHEET_TYPES = /\.(xlsx|xlsm|xls|csv)$/i;
// Claude's Python sandbox (code execution) with Anthropic's ready-made Excel skill. It has no internet.
const CODE_TOOL = { type: 'code_execution_20250825', name: 'code_execution' };
const SKILLS = [{ type: 'anthropic', skill_id: 'xlsx', version: 'latest' }];
const SANDBOX_NOTE = `

PYTHON SANDBOX
You also have code_execution: a Python sandbox (pandas, openpyxl) with Anthropic's Excel skill. When a spreadsheet is loaded, a copy of the file exactly as uploaded is put in the sandbox (find it with: find / -maxdepth 4 -type f \\( -iname '*.xlsx' -o -iname '*.xlsm' -o -iname '*.xls' -o -iname '*.csv' \\) -mmin -600 2>/dev/null | grep -v -e /usr -e /opt -e /skills). It stays there for the rest of the conversation, with anything you save next to it.
Use it to understand the file properly: every tab, merged cells, formulas (openpyxl data_only=False) and their saved values (data_only=True), tables side by side or stacked, headings over two rows, totals typed in, numbers typed as text. Sheets are typed by hand and every school's is different, so look before you assume.
The sandbox cannot see or change the page, and the page cannot see the sandbox. To make the page use what you found, call the page tools (set_header_row, set_column, set_cells, enter_rows, use_all_sheets, choose_sheet). Edits made on the page are not in the sandbox copy.
The page's engine is the source of truth for judgements, phase totals, limits and simulations. Use Python to read and to check; if your figures disagree with the page, find out why and fix how the page reads the sheet. Never report your own Python judgements instead of the page's.
Print compact results only (no whole sheets). Do not mention Python, code or the sandbox to the person unless they ask; say you checked the file.`;

// Container ids are signed so a visitor can only continue a sandbox this site gave them.
async function sign(env, id) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode('gis-container:' + env.ANTHROPIC_API_KEY), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const mac = new Uint8Array(await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(id)));
  return btoa(String.fromCharCode(...mac.slice(0, 18))).replace(/\+/g, '-').replace(/\//g, '_');
}
async function openContainer(env, token) {
  if (typeof token !== 'string' || token.length > 200) return null;
  const i = token.lastIndexOf('.'); if (i < 1) return null;
  const id = token.slice(0, i);
  return /^[A-Za-z0-9_-]{4,150}$/.test(id) && (await sign(env, id)) === token.slice(i + 1) ? id : null;
}
function apiHeaders(env, extra) { return Object.assign({ 'x-api-key': env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' }, extra || {}); }
// Long runs (max effort, code in the sandbox) can go minutes without a byte, and an idle request is cut off
// (HTTP 524). So Claude's reply is streamed and put back together here into one message.
async function claudeStream(env, payload) {
  const res = await fetch(base(env) + '/messages', { method: 'POST', headers: apiHeaders(env, { 'content-type': 'application/json' }), body: JSON.stringify(Object.assign({}, payload, { stream: true })) });
  if (!res.ok) { const raw = await res.text(); let data = {}; try { data = JSON.parse(raw); } catch { /* not JSON */ } return { ok: false, status: res.status, data, raw }; }
  const msg = { content: [], stop_reason: null, container: null }, partial = {};
  let err = null, buf = '';
  const reader = res.body.getReader(), dec = new TextDecoder();
  const handle = (ev) => {
    const b = ev.index != null ? msg.content[ev.index] : null, d = ev.delta || {};
    if (ev.type === 'message_start') { if (ev.message && ev.message.container) msg.container = ev.message.container; }
    else if (ev.type === 'content_block_start') { msg.content[ev.index] = ev.content_block; if (/tool_use$/.test(ev.content_block.type)) partial[ev.index] = ''; }
    else if (ev.type === 'content_block_delta' && b) {
      if (d.type === 'text_delta') b.text = (b.text || '') + d.text;
      else if (d.type === 'thinking_delta') b.thinking = (b.thinking || '') + d.thinking;
      else if (d.type === 'signature_delta') b.signature = (b.signature || '') + d.signature;
      else if (d.type === 'input_json_delta') partial[ev.index] = (partial[ev.index] || '') + d.partial_json;
      else if (d.type === 'citations_delta') (b.citations = b.citations || []).push(d.citation);
    } else if (ev.type === 'content_block_stop' && partial[ev.index] != null) {
      try { b.input = partial[ev.index] ? JSON.parse(partial[ev.index]) : {}; } catch { b.input = {}; }
      delete partial[ev.index];
    } else if (ev.type === 'message_delta') { if (d.stop_reason) msg.stop_reason = d.stop_reason; if (d.container) msg.container = d.container; if (ev.container) msg.container = ev.container; }
    else if (ev.type === 'error') err = ev.error || { message: 'stream error' };
  };
  for (;;) {
    const { done, value } = await reader.read();
    if (value) buf += dec.decode(value, { stream: true });
    let i;
    while ((i = buf.indexOf('\n\n')) >= 0) {
      const chunk = buf.slice(0, i); buf = buf.slice(i + 2);
      const data = chunk.split('\n').filter((l) => l.startsWith('data:')).map((l) => l.slice(5).trim()).join('');
      if (data) { try { handle(JSON.parse(data)); } catch { /* skip a bad line */ } }
    }
    if (done) break;
  }
  msg.content = msg.content.filter(Boolean);
  return err ? { ok: false, status: 500, data: { error: err }, raw: '' } : { ok: true, status: 200, data: msg };
}
// Sends spaces while the work runs so the visitor's connection stays open, then the JSON (leading spaces are valid JSON).
function keepAlive(ctx, work) {
  const { readable, writable } = new TransformStream(), w = writable.getWriter(), enc = new TextEncoder();
  const tick = setInterval(() => { w.write(enc.encode(' ')).catch(() => {}); }, 10000);
  ctx.waitUntil((async () => {
    let out;
    try { out = await work(); } catch (e) { out = { _status: 502, error: 'The assistant stopped unexpectedly. Try again.' }; }
    clearInterval(tick);
    try { await w.write(enc.encode(JSON.stringify(out))); await w.close(); } catch { /* visitor left */ }
  })());
  return new Response(readable, { headers: { 'content-type': 'application/json', 'cache-control': 'no-store' } });
}
function base(env) { return env.ANTHROPIC_BASE || API; } // ANTHROPIC_BASE is only for local tests
async function uploadSheet(env, up) {
  const name = clean(up && up.name, 120).replace(/[<>:"|?*\\/]/g, '_');
  if (!SHEET_TYPES.test(name) || typeof up.data !== 'string') return null;
  let bytes;
  try { bytes = Uint8Array.from(atob(up.data), (c) => c.charCodeAt(0)); } catch { return null; }
  if (!bytes.length || bytes.length > 9 * 1024 * 1024) return null;
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  form.append('expires_in_seconds', '3600'); // deleted straight after use; this is the safety net
  const res = await fetch(base(env) + '/files', { method: 'POST', headers: apiHeaders(env), body: form });
  if (!res.ok) return null;
  const f = await res.json().catch(() => ({}));
  return f.id || null;
}
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
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname === '/api/health') return json({ ok: Boolean(env.ANTHROPIC_API_KEY), reports: Boolean(env.GIS_GITHUB_TOKEN), pdf: Boolean(env.BROWSER) });

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

    if (url.pathname === '/api/pdf') {
      // Prints the report page (built in the visitor's browser) to PDF with Cloudflare Browser Run. Nothing is stored.
      if (request.method !== 'POST') return json({ error: 'Use POST.' }, 405);
      if (!sameSite(request, url)) return json({ error: 'Requests are only accepted from this site.' }, 403);
      if (!env.BROWSER) return json({ error: 'PDF printing is not set up.' }, 503);
      if (Number(request.headers.get('content-length') || 0) > 8 * 1024 * 1024) return json({ error: 'That report is too large.' }, 413);
      if (!(await allowed(env.PDF_LIMIT, request.headers.get('cf-connecting-ip') || 'unknown'))) return json({ error: 'Too many reports right now. Wait a minute and try again.' }, 429);
      const html = await request.text();
      if (!/^<!doctype html>/i.test(html.trim())) return json({ error: 'Not a report.' }, 400);
      let browser;
      try {
        browser = await puppeteer.launch(env.BROWSER);
        const page = await browser.newPage();
        await page.setJavaScriptEnabled(false);
        await page.setRequestInterception(true);
        page.on('request', (r) => {
          const u = r.url();
          if (u.startsWith('data:') || u.startsWith('about:') || /^https:\/\/fonts\.(googleapis|gstatic)\.com\//.test(u)) r.continue(); else r.abort();
        });
        await page.setContent(html, { waitUntil: 'networkidle0', timeout: 25000 });
        const pdf = await page.pdf({ preferCSSPageSize: true, printBackground: true });
        return new Response(pdf, { headers: { 'content-type': 'application/pdf', 'cache-control': 'no-store' } });
      } catch (e) {
        return json({ error: 'The PDF could not be made right now.' }, 502);
      } finally {
        if (browser) { try { await browser.close(); } catch { /* already closed */ } }
      }
    }

    if (url.pathname === '/api/gsheet') {
      // Fetches a Google Sheet shared as "Anyone with the link" as .xlsx for the page. Not stored.
      if (request.method !== 'GET') return json({ error: 'Use GET.' }, 405);
      if (!sameSite(request, url)) return json({ error: 'Requests are only accepted from this site.' }, 403);
      const id = url.searchParams.get('id') || '';
      if (!/^[A-Za-z0-9_-]{20,100}$/.test(id)) return json({ error: 'That is not a Google Sheets link.' }, 400);
      if (!(await allowed(env.CHAT_LIMIT, request.headers.get('cf-connecting-ip') || 'unknown'))) return json({ error: 'Too many requests right now.' }, 429);
      const res = await fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx', { redirect: 'follow' });
      const type = res.headers.get('content-type') || '';
      if (!res.ok || /text\/html/.test(type)) return json({ error: 'private' }, res.status === 404 ? 404 : 403);
      const buf = await res.arrayBuffer();
      if (buf.byteLength > 15 * 1024 * 1024) return json({ error: 'That sheet is too large.' }, 413);
      return new Response(buf, { headers: { 'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'cache-control': 'no-store' } });
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

      // The spreadsheet goes into Claude's sandbox for this request only; the uploaded copy is deleted right after.
      let fileId = null;
      if (body.upload) {
        fileId = await uploadSheet(env, body.upload);
        const last = messages[messages.length - 1];
        if (fileId && last && last.role === 'user') {
          const parts = typeof last.content === 'string' ? [{ type: 'text', text: last.content }] : Array.isArray(last.content) ? last.content.slice() : [];
          parts.push({ type: 'container_upload', file_id: fileId });
          messages[messages.length - 1] = { role: 'user', content: parts };
        }
      }
      const containerId = await openContainer(env, body.container);
      const payload = (cid) => ({ model: env.MODEL || 'claude-haiku-5-5', max_tokens: 32000, output_config: { effort: env.EFFORT || 'max' }, system: PROMPT + SANDBOX_NOTE, messages, tools: tools.concat([CODE_TOOL]), container: cid ? { id: cid, skills: SKILLS } : { skills: SKILLS } });
      return keepAlive(ctx, async () => {
        const t0 = Date.now(); let r, failed = '';
        try { r = await claudeStream(env, payload(containerId)); }
        catch (e) { failed = String((e && e.message) || e).slice(0, 200); }
        finally { if (fileId) ctx.waitUntil(fetch(base(env) + '/files/' + fileId, { method: 'DELETE', headers: apiHeaders(env) }).catch(() => {})); }
        const secs = Math.round((Date.now() - t0) / 1000);
        if (failed) return { _status: 502, error: 'Claude could not be reached (' + failed + ', after ' + secs + 's). Try again.' };
        if (!r.ok) {
          const m = (r.data.error && r.data.error.message) || ('HTTP ' + r.status + ' after ' + secs + 's: ' + clean(r.raw, 200));
          // A sandbox that expired (they last 30 days) or was lost: the page starts a new one and sends the file again.
          if (containerId && /container/i.test(m)) return { _status: 409, error: 'container_expired' };
          return { _status: 502, error: m };
        }
        const cid = r.data.container && r.data.container.id;
        return { content: r.data.content, stop_reason: r.data.stop_reason, container: cid ? cid + '.' + (await sign(env, cid)) : null, uploaded: !!fileId };
      });
    }

    return env.ASSETS.fetch(request);
  }
};
