  /* ---------- the assistant (runs inside the app scope) ---------- */
  var AS = { turns: [], log: [], busy: false, ctl: null, mode: null };
  var openAssistantWith = null;
  var INSTRUCTIONS = window.__ASSISTANT_PROMPT__ || '';
  function snapshot() { return JSON.parse(JSON.stringify({ cfg: st.cfg, edits: st.edits, sheetName: st.sheetName, headerRow: st.headerRow || null, sel: st.sel })); }
  function restore(s) { st.cfg = s.cfg; st.edits = s.edits; st.headerRow = s.headerRow; st.sel = s.sel; if (s.sheetName !== st.sheetName && st.src.getWorksheet(s.sheetName)) { st.sheetName = s.sheetName; } st.grid = gridFromWs(st.src.getWorksheet(st.sheetName)); recompute(); }
  function record(label, before, extra) { AS.log.unshift({ id: Date.now() + Math.random(), label: label, before: before, extra: extra || null, when: new Date() }); renderLog(); }
  function letterToCol(s) { s = String(s || '').trim().toUpperCase(); if (!/^[A-Z]{1,3}$/.test(s)) return null; var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function overview() {
    var o = { loaded: !!st.grid, file: st.srcName || null };
    if (!st.src) return o;
    o.sheets = st.src.worksheets.map(function (w) { return w.name; }); o.sheet = st.sheetName; o.headerRowOverride = st.headerRow || null;
    o.sheetSize = st.grid ? { rows: st.grid.maxR, cols: E.colLetter(st.grid.maxC) } : null;
    if (st.parsed) o.tables = st.parsed.blocks.map(function (b) {
      var cols = {}; Object.keys(b.col).forEach(function (k) { if (b.col[k] != null) cols[E.ROLE_LABEL[k]] = E.colLetter(b.col[k]) + ' (' + String(st.grid.cells[b.headerRow + ',' + b.col[k]] || '').slice(0, 50) + ')'; });
      return { title: b.title, headerRow: b.headerRow, lastRow: b.endRow, columns: cols, gradeRows: b.rows.length, totalRows: b.totals.length };
    });
    var c = st.cfg;
    o.rules = { acceptableFromPct: c.acc * 100, goodFromPct: c.good * 100, veryGoodFromPct: c.vg * 100, outstandingFromPct: c.out * 100, veryWeakBelowPct: c.vw == null ? null : c.vw * 100,
      weakLimitPct: c.capWeak * 100, outstandingLimitPct: c.capOut * 100, phaseTotalMethod: c.totalMethod, rounding: c.rounding, counts: c.basis, useAssessed: c.useAssessed, points: c.points, simulatedSheetTitle: c.simTitle };
    if (st.base) {
      o.units = st.base.units.map(function (u) { return { table: u.block, group: u.group, phase: u.phase, grades: u.grades.map(function (g) { return { grade: g.grade, row: g.r, students: g.E, assessed: g.A, expectedOrBetter: g.G, betterThanExpected: g.H, judgement: g.judg }; }), total: u.total ? { judgement: u.total.judg, average: u.total.avg, roundingToConfirm: u.total.roundingFlag } : null }; });
      var k = E.check(E.counted(st.base.items, c.basis), c);
      o.actual = { counted: k.n, weak: k.weak, weakAllowed: k.maxW, outstanding: k.out, outstandingAllowed: k.maxO };
      o.fixes = E.findFixes(st.parsed, c).fixes.map(function (f) { return f.what + ': ' + f.detail; });
      if (st.sim) o.simulation = !st.sim.needed ? 'not needed' : st.sim.failed ? 'not possible with these rules' : st.sim.options.map(function (op, i) { return { option: String.fromCharCode(65 + i), selected: i === st.sel, students: op.students, judgementChanges: op.changed.length, largestChangePoints: +op.maxPts.toFixed(1), grades: op.cells.map(function (x) { return x.to.grade + ' +' + x.dG; }) }; });
    } else if (st.grid) o.problem = 'No progress table was recognised with the current settings.';
    o.step = st.step;
    return o;
  }
  var TOOLS = [
    { name: 'get_overview', description: 'Returns the current state as JSON: file, sheets, detected tables and column mapping, rules, every grade row, results, corrections and simulation options. Call this first, and again after changes to check them.',
      inputSchema: { type: 'object', properties: {} }, execute: function () { return overview(); } },
    { name: 'read_cells', description: 'Reads raw cell values from the loaded sheet (as uploaded, including any fixes made with set_cell). Use it to understand an unusual layout. Returns {cell: value} for non-empty cells. Range like "A1:L30", at most 600 cells.',
      inputSchema: { type: 'object', properties: { range: { type: 'string' }, sheet: { type: 'string' } }, required: ['range'] },
      execute: function (i) {
        var ws = st.src && st.src.getWorksheet(i.sheet ? String(i.sheet) : st.sheetName); if (!ws) throw new Error('No such sheet');
        var m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(String(i.range).trim()); if (!m) throw new Error('Use a range like A1:K20');
        var c1 = letterToCol(m[1]), r1 = +m[2], c2 = letterToCol(m[3]), r2 = +m[4];
        if ((c2 - c1 + 1) * (r2 - r1 + 1) > 600) throw new Error('Range too large, ask for 600 cells or fewer');
        var g = gridFromWs(ws), out = {};
        for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) { var v = g.cells[r + ',' + c]; if (v != null) out[E.colLetter(c) + r] = v; }
        return out;
      } },
    { name: 'choose_sheet', description: 'Switches to another sheet (tab) of the uploaded workbook.', inputSchema: { type: 'object', properties: { name: { type: 'string' } }, required: ['name'] },
      execute: function (i) { var n = String(i.name); if (!st.src.getWorksheet(n)) throw new Error('No sheet called ' + n); var b = snapshot(); st.headerRow = null; pickSheet(n); record('Switched to the sheet “' + n + '”', b); return overview().tables || []; } },
    { name: 'set_header_row', description: 'Tells the page which row holds the column headings, when it was not found automatically. Pass row 0 to go back to automatic.', inputSchema: { type: 'object', properties: { row: { type: 'integer' } }, required: ['row'] },
      execute: function (i) { var b = snapshot(); st.headerRow = +i.row || null; st.cfg.columns = {}; st.edits = {}; recompute(); record(st.headerRow ? 'Headings read from row ' + st.headerRow : 'Headings found automatically', b); return overview().tables || []; } },
    { name: 'set_column', description: 'Sets which column holds a role. Roles: phase, group, grade, E (number of students), F (% assessed), G (made expected or better), H (made better than expected), I (% expected or better), J (% better than expected), K (judgement). Column is a letter like "D", or "" for not used.',
      inputSchema: { type: 'object', properties: { role: { type: 'string', enum: E.ROLES }, column: { type: 'string' } }, required: ['role', 'column'] },
      execute: function (i) { var role = String(i.role); if (E.ROLES.indexOf(role) < 0) throw new Error('Unknown role'); var col = i.column === '' ? '' : letterToCol(i.column); if (col === null) throw new Error('Column must be a letter'); var b = snapshot(); st.cfg.columns[role] = col; st.edits = {}; recompute(); record('“' + E.ROLE_LABEL[role] + '” read from column ' + (col === '' ? '(none)' : E.colLetter(col)), b); return overview().tables || []; } },
    { name: 'set_rules', description: 'Changes rules. Percentages are numbers from 0 to 100. Only change the thresholds if the user asks: never change them to make results fit. phaseTotalMethod: "points" or "counts". rounding: "unspecified", "nearest", "down" or "up". counts: "both", "grades" or "phases". points: object of judgement name to points.',
      inputSchema: { type: 'object', properties: { acceptableFromPct: { type: 'number' }, goodFromPct: { type: 'number' }, veryGoodFromPct: { type: 'number' }, outstandingFromPct: { type: 'number' }, veryWeakBelowPct: { type: ['number', 'null'] }, weakLimitPct: { type: 'number' }, outstandingLimitPct: { type: 'number' }, phaseTotalMethod: { type: 'string' }, rounding: { type: 'string' }, counts: { type: 'string' }, useAssessed: { type: 'boolean' }, substantialPoints: { type: 'number' }, simulatedSheetTitle: { type: 'string' }, points: { type: 'object' } } },
      execute: function (i) {
        var b = snapshot(), c = st.cfg, done = [];
        function p(k, key, label) { if (i[k] !== undefined) { c[key] = i[k] === null ? null : Math.max(0, Math.min(100, +i[k])) / 100; done.push(label + ' ' + (i[k] === null ? 'off' : i[k] + '%')); } }
        p('acceptableFromPct', 'acc', 'Acceptable from'); p('goodFromPct', 'good', 'Good from'); p('veryGoodFromPct', 'vg', 'Very Good from'); p('outstandingFromPct', 'out', 'Outstanding from'); p('veryWeakBelowPct', 'vw', 'Very Weak below');
        p('weakLimitPct', 'capWeak', 'Weak limit'); p('outstandingLimitPct', 'capOut', 'Outstanding limit');
        if (i.phaseTotalMethod === 'points' || i.phaseTotalMethod === 'counts') { c.totalMethod = i.phaseTotalMethod; done.push('Phase totals ' + (c.totalMethod === 'points' ? 'from the average of grades' : 'from total students')); }
        if (['unspecified', 'nearest', 'down', 'up'].indexOf(i.rounding) >= 0) { c.rounding = i.rounding; done.push('Rounding: ' + i.rounding); }
        if (['both', 'grades', 'phases'].indexOf(i.counts) >= 0) { c.basis = i.counts; done.push('Counting ' + i.counts); }
        if (typeof i.useAssessed === 'boolean') { c.useAssessed = i.useAssessed; done.push((i.useAssessed ? 'Using' : 'Not using') + ' the % assessed column'); }
        if (i.substantialPoints != null) { c.substantialPts = +i.substantialPoints; done.push('Substantial above ' + i.substantialPoints + ' points'); }
        if (i.simulatedSheetTitle) { c.simTitle = String(i.simulatedSheetTitle).slice(0, 80); done.push('Simulated sheet title changed'); }
        if (i.points && typeof i.points === 'object') { Object.keys(i.points).forEach(function (k) { if (E.CATS.indexOf(k) >= 0) c.points[k] = i.points[k] === null ? null : +i.points[k]; }); done.push('Points changed'); }
        if (!done.length) throw new Error('Nothing to change');
        st.sel = 0; recompute(); record(done.join('. '), b); return { changed: done, actual: overview().actual };
      } },
    { name: 'edit_number', description: 'Corrects one number for a grade on the page (the uploaded file itself is kept). field: "students", "expected" (made expected or better) or "better" (made better than expected). Identify the grade by its row number from get_overview.',
      inputSchema: { type: 'object', properties: { row: { type: 'integer' }, field: { type: 'string', enum: ['students', 'expected', 'better'] }, value: { type: 'integer' } }, required: ['row', 'field', 'value'] },
      execute: function (i) {
        var it = st.base && st.base.items.filter(function (x) { return x.kind === 'grade' && x.r === +i.row; })[0]; if (!it) throw new Error('No grade on row ' + i.row);
        var k = { students: 'E', expected: 'G', better: 'H' }[i.field]; if (!k) throw new Error('Unknown field');
        var b = snapshot(), old = it[k]; (st.edits[it.id] = st.edits[it.id] || {})[k] = Math.max(0, Math.round(+i.value)); recompute();
        record(it.grade + ': ' + { E: 'students', G: 'expected or better', H: 'better than expected' }[k] + ' ' + old + ' → ' + Math.round(+i.value), b); return 'Done';
      } },
    { name: 'set_cell', description: 'Changes one cell in the working copy of the uploaded sheet, for layout problems (a missing heading, a phase label, a "Total" label, a number stored as text). The original upload is kept untouched in the workbook. Value may be text, a number, or null to clear.',
      inputSchema: { type: 'object', properties: { cell: { type: 'string' }, value: { type: ['string', 'number', 'null'] } }, required: ['cell', 'value'] },
      execute: function (i) {
        var m = /^([A-Z]+)(\d+)$/i.exec(String(i.cell).trim()); if (!m) throw new Error('Cell like D5');
        var ws = st.src.getWorksheet(st.sheetName), cell = ws.getCell(String(i.cell).toUpperCase()), prev = cell.value, b = snapshot();
        var v = i.value; if (typeof v === 'string' && v.trim() !== '' && !isNaN(+v)) v = +v;
        cell.value = v; st.grid = gridFromWs(ws); recompute();
        record('Cell ' + String(i.cell).toUpperCase() + ' set to “' + (v == null ? '(empty)' : v) + '”', b, { cell: String(i.cell).toUpperCase(), prev: prev, sheet: st.sheetName }); return overview().tables || [];
      } },
    { name: 'choose_option', description: 'Selects a simulation option by letter (A is the smallest).', inputSchema: { type: 'object', properties: { letter: { type: 'string' } }, required: ['letter'] },
      execute: function (i) { var n = String(i.letter).toUpperCase().charCodeAt(0) - 65; if (!st.sim || !st.sim.options || !st.sim.options[n]) throw new Error('No such option'); var b = snapshot(); st.sel = n; st.wb = null; render(); record('Chose simulation option ' + String.fromCharCode(65 + n), b); return 'Done'; } },
    { name: 'go_to_step', description: 'Shows a step of the page so the user can see the result: 1 Load, 2 Check, 3 Rules, 4 Results, 5 Simulation, 6 Download.', inputSchema: { type: 'object', properties: { step: { type: 'integer' } }, required: ['step'] },
      execute: function (i) { go(+i.step); return 'Showing step ' + st.step; } }
  ];

  /* ---------- transports ---------- */
  function detect() {
    if (window.claude && typeof window.claude.use === 'function') {
      return window.claude.use('sample').then(function (s) {
        if (!s) return null;
        return (s.limits ? s.limits().catch(function () { return null; }) : Promise.resolve(null)).then(function (l) { if (l && l.tools) { AS.sample = s; return 'claude'; } return null; });
      }).catch(function () { return null; });
    }
    return fetch('/api/health', { method: 'GET' }).then(function (r) { return r.ok ? r.json() : null; }).then(function (j) { return j && j.ok ? 'api' : null; }).catch(function () { return null; });
  }
  function runClaude(userText, onText) {
    var turns = [{ role: 'user', content: INSTRUCTIONS + '\n\nCURRENT STATE (JSON):\n' + JSON.stringify(overview()).slice(0, 60000) }].concat(AS.turns.slice(-12)).concat([{ role: 'user', content: userText }]);
    return AS.sample(turns, { tools: TOOLS.map(function (t) { return { name: t.name, description: t.description, inputSchema: t.inputSchema, execute: function (inp) { activity(t.name, inp); return t.execute(inp || {}); } }; }), cache: false, signal: AS.ctl.signal, onText: function (u) { onText(u.text); } })
      .then(function (r) { return r.text; });
  }
  function runApi(userText, onText) {
    var msgs = AS.turns.slice(-12).map(function (t) { return { role: t.role, content: t.content }; });
    msgs.push({ role: 'user', content: 'CURRENT STATE (JSON):\n' + JSON.stringify(overview()).slice(0, 60000) + '\n\nREQUEST:\n' + userText });
    var tools = TOOLS.map(function (t) { return { name: t.name, description: t.description, input_schema: t.inputSchema }; });
    var said = '';
    function round(n) {
      if (n > 8) return Promise.resolve(said || 'I stopped after several steps. Check the changes listed below.');
      return fetch('/api/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ messages: msgs, tools: tools }), signal: AS.ctl.signal })
        .then(function (r) { if (!r.ok) return r.json().catch(function () { return {}; }).then(function (j) { throw { code: 'api', message: j.error || ('The assistant is unavailable (' + r.status + ').') }; }); return r.json(); })
        .then(function (res) {
          var blocks = res.content || [], text = blocks.filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('\n');
          if (text) { said += (said ? '\n\n' : '') + text; onText(said); }
          msgs.push({ role: 'assistant', content: blocks });
          var uses = blocks.filter(function (b) { return b.type === 'tool_use'; });
          if (res.stop_reason !== 'tool_use' || !uses.length) return said;
          var results = uses.map(function (u) {
            var t = TOOLS.filter(function (x) { return x.name === u.name; })[0], out, err = false;
            try { if (!t) throw new Error('Unknown tool'); activity(u.name, u.input); out = t.execute(u.input || {}); } catch (e) { out = 'Error: ' + (e && e.message || e); err = true; }
            return { type: 'tool_result', tool_use_id: u.id, content: typeof out === 'string' ? out : JSON.stringify(out).slice(0, 30000), is_error: err };
          });
          msgs.push({ role: 'user', content: results });
          return round(n + 1);
        });
    }
    return round(1);
  }

  /* ---------- drawer UI ---------- */
  var drawer = $('asst'), log = $('asst-log'), input = $('asst-in');
  function openDrawer(open) { drawer.dataset.open = open ? 'true' : 'false'; drawer.setAttribute('aria-hidden', open ? 'false' : 'true'); document.body.classList.toggle('asst-open', open); if (open) setTimeout(function () { input.focus(); }, 60); }
  $('asst-open').addEventListener('click', function () { openDrawer(drawer.dataset.open !== 'true'); });
  $('asst-close').addEventListener('click', function () { openDrawer(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && drawer.dataset.open === 'true') openDrawer(false); });
  function bubble(role, text, cls) { var d = document.createElement('div'); d.className = 'bub ' + role + (cls ? ' ' + cls : ''); d.textContent = text; log.appendChild(d); log.scrollTop = log.scrollHeight; return d; }
  var NICE = { get_overview: 'Looking at the sheet', read_cells: 'Reading cells', choose_sheet: 'Switching sheet', set_header_row: 'Setting the heading row', set_column: 'Matching a column', set_rules: 'Changing rules', edit_number: 'Correcting a number', set_cell: 'Fixing a cell', choose_option: 'Choosing an option', go_to_step: 'Opening a step' };
  var actEl = null;
  function activity(name) { if (actEl) actEl.textContent = (NICE[name] || name) + '…'; }
  function renderLog() {
    var box = $('asst-changes');
    box.hidden = !AS.log.length;
    $('asst-changes-list').innerHTML = AS.log.map(function (x, i) { return '<li><span class="dot sim" aria-hidden="true"></span><span>' + esc(x.label) + '</span>' + (i === 0 ? '<button type="button" class="btn quiet" data-undo="' + x.id + '">Undo</button>' : '') + '</li>'; }).join('');
    var banner = $('asst-banner');
    if (banner) { banner.hidden = !AS.log.length; $('asst-banner-t').textContent = AS.log.length ? 'The assistant made ' + AS.log.length + ' change' + (AS.log.length === 1 ? '' : 's') + '. Latest: ' + AS.log[0].label + '.' : ''; }
  }
  document.addEventListener('click', function (e) {
    var u = e.target.closest && e.target.closest('[data-undo]'); if (!u) return;
    var x = AS.log[0]; if (!x) return;
    if (x.extra && x.extra.cell) { var ws = st.src.getWorksheet(x.extra.sheet); if (ws) ws.getCell(x.extra.cell).value = x.extra.prev; }
    AS.log.shift(); restore(x.before); renderLog(); bubble('note', 'Undone: ' + x.label + '.');
  });
  $('asst-banner-open') && $('asst-banner-open').addEventListener('click', function () { openDrawer(true); });
  function send(text) {
    text = String(text || '').trim(); if (!text || AS.busy) return;
    if (!AS.mode) { bubble('note', 'The assistant is not available here.'); return; }
    AS.busy = true; AS.ctl = new AbortController(); $('asst-send').hidden = true; $('asst-stop').hidden = false;
    bubble('user', text); input.value = '';
    var out = bubble('bot', ''); actEl = document.createElement('div'); actEl.className = 'act'; actEl.textContent = 'Thinking…'; out.appendChild(actEl);
    var p = AS.mode === 'claude' ? runClaude(text, function (t) { out.textContent = t; out.appendChild(actEl); }) : runApi(text, function (t) { out.textContent = t; out.appendChild(actEl); });
    p.then(function (final) {
      out.textContent = final || 'Done.'; AS.turns.push({ role: 'user', content: text }, { role: 'assistant', content: final || 'Done.' });
    }, function (e) {
      var c = e && e.code;
      if (c === 'cancelled' || (e && e.name === 'AbortError')) out.textContent = (e.text || out.textContent.replace(/\S*…$/, '')) + ' (stopped)';
      else if (c === 'not_granted' || c === 'sampling_disabled') { out.textContent = 'The assistant needs permission to use Claude, and it was not given.'; }
      else if (c === 'rate_limited') out.textContent = 'Too many requests right now. Try again in a little while.';
      else out.textContent = (e && e.message) || 'Something went wrong. Try again.';
      out.classList.add('err');
    }).then(function () { AS.busy = false; $('asst-send').hidden = false; $('asst-stop').hidden = true; actEl = null; input.focus(); });
  }
  openAssistantWith = function (q) { openDrawer(true); send(q); };
  $('asst-form').addEventListener('submit', function (e) { e.preventDefault(); send(input.value); });
  input.addEventListener('keydown', function (e) { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(input.value); } });
  $('asst-stop').addEventListener('click', function () { if (AS.ctl) AS.ctl.abort(); });
  $('asst-sugg').addEventListener('click', function (e) { var b = e.target.closest('[data-q]'); if (b) send(b.dataset.q); });
  detect().then(function (m) {
    AS.mode = m;
    $('asst-open').hidden = !m;
    $('asst-where').textContent = m === 'claude' ? 'Uses Claude on your own Claude account. The parts of your sheet it needs are sent to Claude when you ask something.' : 'Uses Claude through this site. The parts of your sheet it needs are sent to Claude when you ask something, and are not stored.';
  });
