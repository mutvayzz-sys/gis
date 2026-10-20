(function () {
  var E = Engine;
  var $ = function (id) { return document.getElementById(id); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  var KEY = { 'Very Weak': 'vw', 'Weak': 'weak', 'Acceptable': 'acc', 'Good': 'good', 'Very Good': 'vg', 'Outstanding': 'out' };
  function dot(j) { return '<span class="dot k-' + (KEY[j] || 'acc') + '" aria-hidden="true"></span>'; }
  function tag(j) { return '<span class="judg">' + dot(j) + esc(j) + '</span>'; }
  function pc(v, d) { return v == null ? '–' : (v * 100).toFixed(d == null ? 1 : d) + '%'; }

  // All state lives here, in memory only.
  var st = { src: null, origWb: null, headerRow: null, srcName: '', sheetName: null, grid: null, parsed: null, cfg: E.defaults(), edits: {}, manual: null, base: null, sim: null, sel: 0, step: 1, wb: null, tabA: 0, tabS: 0 };

  /* ---------- reading files ---------- */
  function gridFromWs(ws) {
    var cells = {}, maxR = 0, maxC = 0, hasF = false;
    ws.eachRow({ includeEmpty: false }, function (row, r) {
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        if (cell.isMerged && cell.master !== cell) return;
        var v = cell.value;
        if (v && typeof v === 'object') {
          if ('formula' in v || 'sharedFormula' in v) { hasF = true; v = v.result; }
          else if (v.richText) v = v.richText.map(function (t) { return t.text; }).join('');
          else if (v instanceof Date) v = v.toISOString().slice(0, 10);
          else if ('text' in v) v = v.text; else if ('error' in v) v = null;
        }
        if (v == null || v === '') return;
        cells[r + ',' + c] = v; maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
      });
    });
    return { cells: cells, maxR: maxR, maxC: maxC, hasFormulas: hasF };
  }
  function aoaToWorkbook(sheets) {
    var wb = new ExcelJS.Workbook();
    sheets.forEach(function (s) { var ws = wb.addWorksheet(s.name.slice(0, 31) || 'Sheet'); s.aoa.forEach(function (row, r) { (row || []).forEach(function (v, c) { if (v != null && v !== '') ws.getCell(r + 1, c + 1).value = v; }); }); });
    return wb;
  }
  function msg(id, t, k) { var m = $(id); if (!m) return; m.textContent = t || ''; m.className = 'msg' + (k ? ' ' + k : ''); }
  function useWorkbook(wb, name, orig) {
    st.src = wb; st.origWb = orig || null; st.headerRow = null; st.srcName = name; st.edits = {}; st.manual = null; st.sel = 0; st.wb = null;
    var withTables = wb.worksheets.filter(function (ws) { return E.parse(gridFromWs(ws), { sheetName: ws.name }).blocks.length; });
    if (!withTables.length) {
      var first = wb.worksheets.filter(function (ws) { return gridFromWs(ws).maxR > 0; })[0];
      if (!first) { msg('msg1', 'This file has no data in it.', 'bad'); st.src = null; return; }
      $('sheet-pick').hidden = wb.worksheets.length < 2;
      $('sheet-list').innerHTML = wb.worksheets.map(function (ws) { return '<button type="button" class="btn' + (ws === first ? ' primary' : '') + '" data-sheet="' + esc(ws.name) + '">' + esc(ws.name) + '</button>'; }).join('');
      pickSheet(first.name);
      msg('msg1', 'Loaded ' + name + ', but the table needs a little help to be read.', '');
      go(2); return;
    }
    if (withTables.length > 1) {
      $('sheet-pick').hidden = false;
      $('sheet-list').innerHTML = withTables.map(function (ws, i) { return '<button type="button" class="btn' + (i === 0 ? ' primary' : '') + '" data-sheet="' + esc(ws.name) + '">' + esc(ws.name) + '</button>'; }).join('');
    } else $('sheet-pick').hidden = true;
    pickSheet(withTables[0].name);
    msg('msg1', 'Loaded ' + name + '.', 'ok');
    go(2);
  }
  function pickSheet(name) {
    st.sheetName = name; st.headerRow = null; st.grid = gridFromWs(st.src.getWorksheet(name)); st.cfg.columns = {}; st.edits = {}; st.sel = 0;
    Array.prototype.forEach.call(document.querySelectorAll('[data-sheet]'), function (b) { b.classList.toggle('primary', b.dataset.sheet === name); });
    recompute();
  }
  function readFile(f) {
    if (!f) return;
    var ext = (f.name.split('.').pop() || '').toLowerCase();
    var rd = new FileReader();
    rd.onerror = function () { msg('msg1', 'That file could not be read. If it came from a download, download it again and check it opens in Excel.', 'bad'); };
    rd.onload = function () {
      var buf = rd.result;
      if (buf.byteLength < 200 && ext === 'xlsx') { msg('msg1', 'This file is almost empty (' + buf.byteLength + ' bytes), so it is not a real spreadsheet. It was probably a failed download. Download it again and check it opens in Excel.', 'bad'); return; }
      if (ext === 'xlsx' || ext === 'xlsm') {
        if (typeof ExcelJS === 'undefined') { msg('msg1', 'The spreadsheet reader did not load. Reload the page, or paste the rows instead.', 'bad'); return; }
        var wb = new ExcelJS.Workbook(), orig = new ExcelJS.Workbook();
        Promise.all([wb.xlsx.load(buf.slice(0)), orig.xlsx.load(buf.slice(0))]).then(function () { useWorkbook(wb, f.name, orig); }, function () { msg('msg1', 'This does not look like a working Excel file. Open it in Excel, save a fresh copy, and try again.', 'bad'); });
      } else {
        try {
          if (typeof XLSX === 'undefined') throw new Error('reader not loaded');
          var x = /^(csv|tsv|txt)$/.test(ext) ? XLSX.read(new TextDecoder().decode(new Uint8Array(buf)), { type: 'string' }) : XLSX.read(new Uint8Array(buf), { type: 'array' });
          var sheets = x.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(x.Sheets[n], { header: 1, raw: true, defval: null }) }; });
          useWorkbook(aoaToWorkbook(sheets), f.name, aoaToWorkbook(sheets));
        } catch (err) { msg('msg1', 'That file could not be read. Try saving it as .xlsx, or paste the rows.', 'bad'); }
      }
    };
    rd.readAsArrayBuffer(f);
  }
  $('file-in').addEventListener('change', function (e) { readFile(e.target.files && e.target.files[0]); e.target.value = ''; });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { readFile(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]); });
  $('sheet-list').addEventListener('click', function (e) { var b = e.target.closest('[data-sheet]'); if (b) pickSheet(b.dataset.sheet); });
  $('open-paste').addEventListener('click', function () { var c = $('paste-card'); c.hidden = !c.hidden; if (!c.hidden) $('paste-box').focus(); });
  $('use-paste').addEventListener('click', function () {
    var t = $('paste-box').value; if (!t.trim()) { msg('msg1', 'Paste some rows first.', 'bad'); return; }
    var lines = t.replace(/\r/g, '').split('\n').filter(function (l) { return l.trim(); });
    var sep = lines.some(function (l) { return l.indexOf('\t') >= 0; }) ? '\t' : ',';
    var aoa = lines.map(function (l) { return l.split(sep).map(function (v) { v = v.trim(); return v === '' ? null : (isNaN(+v) ? v : +v); }); });
    var wb = aoaToWorkbook([{ name: 'Pasted', aoa: aoa }]), orig = aoaToWorkbook([{ name: 'Pasted', aoa: aoa }]);
    if (!E.parse(gridFromWs(wb.worksheets[0]), {}).blocks.length) {
      // plain rows: Grade, Phase, Students, Expected or better, Better than expected
      var rows = aoa.filter(function (r) { return r[0] != null && r[2] != null; }).map(function (r) { return { grade: String(r[0]), phase: r[1], E: +r[2], G: +r[3], H: +r[4] }; });
      if (!rows.length) { msg('msg1', 'Those rows could not be read. Include the header row, or use: Grade, Phase, Students, Expected or better, Better than expected.', 'bad'); return; }
      wb = Book.synthetic(rows, 'Pasted data'); orig = Book.synthetic(rows, 'Pasted data');
    }
    $('paste-box').value = ''; $('paste-card').hidden = true;
    useWorkbook(wb, 'pasted rows', orig);
  });
  $('blank').addEventListener('click', function () {
    st.manual = []; for (var g = 1; g <= 12; g++) st.manual.push({ grade: 'G' + g, phase: g <= 5 ? 2 : g <= 8 ? 3 : 4, E: null, G: null, H: null });
    st.src = null; st.srcName = 'numbers typed in'; st.edits = {}; st.sel = 0;
    manualToSource(); go(2);
  });
  function manualToSource() {
    var rows = st.manual.filter(function (r) { return r.E > 0 && r.G != null && r.H != null; });
    var rs = rows.length ? rows : [{ grade: 'G1', phase: 1, E: 1, G: 0, H: 0 }];
    st.src = Book.synthetic(rs, 'Your numbers'); st.origWb = Book.synthetic(rs, 'Your numbers');
    st.sheetName = 'Data'; st.grid = gridFromWs(st.src.getWorksheet('Data'));
    st.empty = !rows.length;
    recompute();
  }

  /* ---------- the pipeline ---------- */
  var timer = null;
  function soon() { clearTimeout(timer); timer = setTimeout(recompute, 140); }
  function recompute() {
    if (!st.grid) return;
    st.parsed = E.parse(st.grid, { columns: st.cfg.columns, sheetName: st.sheetName, source: st.srcName, headerRow: st.headerRow });
    st.parsed.blocks.forEach(function (b) { b.rows.forEach(function (x) { var e = st.edits[b.id + ':' + x.r]; if (e) Object.keys(e).forEach(function (k) { x[k] = e[k]; }); }); });
    if (!st.parsed.blocks.length) { st.base = null; st.sim = null; render(); return; }
    st.base = E.build(st.parsed, st.cfg);
    st.sim = E.simulate(st.parsed, st.cfg);
    if (st.sim.options && st.sel >= st.sim.options.length) st.sel = 0;
    st.wb = null;
    var subj = st.parsed.blocks[0].title;
    $('title').textContent = subj && subj !== st.sheetName ? subj + ' Progress Judgements' : 'Progress Judgements';
    render();
  }
  function workbook() {
    if (!st.wb && st.base) st.wb = Book.build({ srcWb: st.src, origWb: st.origWb, sheetName: st.sheetName, parsed: st.parsed, cfg: st.cfg, base: st.base, sim: st.sim, sel: st.sel });
    return st.wb;
  }

  /* ---------- steps ---------- */
  function ready() { return !!(st.base && !st.empty); }
  function go(n) {
    n = +n;
    if (n > 1 && !st.grid) return;
    if (n > 3 && !ready()) n = 2;
    st.step = n;
    for (var i = 1; i <= 6; i++) $('step-' + i).hidden = i !== n;
    render();
    var top = document.querySelector('.stepper'); if (top && window.scrollY > top.offsetTop) window.scrollTo({ top: top.offsetTop, behavior: 'auto' });
  }
  document.addEventListener('click', function (e) { var b = e.target.closest && e.target.closest('[data-go]'); if (b && !b.disabled) go(b.dataset.go); });
  function renderStepper() {
    Array.prototype.forEach.call(document.querySelectorAll('#stepper button'), function (b) {
      var n = +b.dataset.go;
      b.disabled = (n > 1 && !st.grid) || (n > 3 && !ready());
      if (n === st.step) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
      b.classList.toggle('done', n < st.step && !b.disabled);
    });
  }
  function render() {
    renderStepper();
    if (st.step === 2) renderCheck();
    if (st.step === 3) renderRules();
    if (st.step === 4 && ready()) renderResults();
    if (st.step === 5 && ready()) renderSim();
    if (st.step === 6 && ready()) renderDownload();
  }

  /* ---------- step 2: check ---------- */
  var ROLE_COLOR = { grade: '#1f3864', E: '#ffd400', F: '#b4b4b9', G: '#0000da', H: '#5656ec', phase: '#8fd19e', group: '#f6b26b', I: '#cfd5ff', J: '#cfd5ff', K: '#d92b2b' };
  var ROLE_SHORT = { phase: 'Phase', group: 'Group', grade: 'Grade', E: 'Students', F: '% assessed', G: 'Expected or better', H: 'Better than expected', I: '% expected or better', J: '% better than expected', K: 'Judgement' };
  function renderCheck() {
    if (st.manual) { renderManual(); return; }
    var p = st.parsed; if (!p) return;
    var nG = 0, nP = 0;
    if (st.base) { nG = st.base.items.filter(function (x) { return x.kind === 'grade'; }).length; nP = st.base.items.length - nG; }
    var nf = !p.blocks.length;
    $('nf-card').hidden = !nf;
    if (nf) { $('nf-ask').hidden = !AS.mode; var hrIn = $('hdr-row'); if (document.activeElement !== hrIn) hrIn.value = st.headerRow || ''; }
    $('colmore').open = nf || $('colmore').open;
    $('found').innerHTML = '<div><b>' + p.blocks.length + '</b><span>table' + (p.blocks.length === 1 ? '' : 's') + (p.blocks.length ? ': ' + esc(p.blocks.map(function (b) { return b.title; }).join(', ')) : '') + '</span></div><div><b>' + nG + '</b><span>grade results</span></div><div><b>' + nP + '</b><span>phase totals</span></div><div><b>' + (st.base ? new Set(st.base.units.map(function (u) { return u.group; })).size : 0) + '</b><span>group' + ((st.base && new Set(st.base.units.map(function (u) { return u.group; })).size === 1) ? '' : 's') + ' of students</span></div>';
    var b = p.blocks[0] || { col: {}, headerRow: st.headerRow || guessHeaderRow() };
    var inv = {}; Object.keys(b.col).forEach(function (k) { if (b.col[k] != null) inv[b.col[k]] = k; });
    var hs = '';
    for (var c = 1; c <= st.grid.maxC; c++) {
      var t = st.grid.cells[b.headerRow + ',' + c]; if (t == null) continue;
      var role = inv[c];
      hs += '<div class="headcell' + (role ? '' : ' none') + '"><span class="L">' + E.colLetter(c) + '</span><span class="t">' + esc(t) + '</span><span class="r">' + (role ? '<span class="role-tag" style="background:' + ROLE_COLOR[role] + '"></span>' + esc(ROLE_SHORT[role]) : 'Not used') + '</span></div>';
    }
    $('headstrip').innerHTML = hs;
    var opts = []; for (var oc = 1; oc <= st.grid.maxC; oc++) { var ot = st.grid.cells[b.headerRow + ',' + oc]; opts.push({ col: oc, label: E.colLetter(oc) + (ot != null ? ': ' + String(ot).replace(/\s+/g, ' ').slice(0, 50) : '') }); }
    var req = { grade: 1, E: 1, G: 1, H: 1 };
    $('colmap').innerHTML = E.ROLES.map(function (r) {
      var cur = b.col[r];
      return '<label class="field"><span><span class="role-tag" style="background:' + ROLE_COLOR[r] + '"></span>' + esc(E.ROLE_LABEL[r]) + (req[r] ? '' : ' (optional)') + '</span><select data-role="' + r + '"><option value="">Not used</option>' + opts.map(function (o) { return '<option value="' + o.col + '"' + (o.col === cur ? ' selected' : '') + '>' + esc(o.label) + '</option>'; }).join('') + '</select></label>';
    }).join('');
    renderEditor();
  }
  function guessHeaderRow() {
    var best = 1, bestN = -1;
    for (var r = 1; r <= Math.min(st.grid.maxR, 30); r++) { var n = 0; for (var c = 1; c <= st.grid.maxC; c++) { var v = st.grid.cells[r + ',' + c]; if (typeof v === 'string' && v.trim()) n++; } if (n > bestN) { bestN = n; best = r; } }
    return best;
  }
  $('hdr-row').addEventListener('change', function () { var v = parseInt(this.value, 10); st.headerRow = v > 0 ? v : null; st.edits = {}; recompute(); });
  $('nf-ask').addEventListener('click', function () { if (typeof openAssistantWith === 'function') openAssistantWith("My sheet wasn't read properly. Please look at it and fix it."); });
  $('colmap').addEventListener('change', function (e) { var s = e.target.closest('select'); if (!s) return; st.cfg.columns[s.dataset.role] = s.value === '' ? '' : +s.value; st.edits = {}; recompute(); });
  function renderEditor() {
    if (!st.base) { $('editor').innerHTML = ''; return; }
    var h = '<thead><tr><th>Grade</th><th>Phase</th><th>Students</th><th>Assessed</th><th>Expected or better</th><th>Better than expected</th><th>Judgement</th></tr></thead><tbody>';
    st.base.units.forEach(function (u) {
      u.grades.forEach(function (g) {
        function inp(k) { return '<input type="number" min="0" step="1" inputmode="numeric" data-edit="' + esc(g.id) + '" data-k="' + k + '" value="' + (g[k] == null ? '' : g[k]) + '" aria-label="' + esc(g.grade + ' ' + k) + '">'; }
        h += '<tr><td>' + esc(g.grade) + '</td><td>' + esc(g.phase) + '</td><td>' + inp('E') + '</td><td>' + (g.A == null ? '' : g.A) + '</td><td>' + inp('G') + '</td><td>' + inp('H') + '</td><td>' + tag(g.judg) + '</td></tr>';
      });
      if (u.total) h += '<tr class="tot"><td>Total</td><td>' + esc(u.phase) + '</td><td>' + u.total.E + '</td><td>' + u.total.A + '</td><td>' + u.total.G + '</td><td>' + u.total.H + '</td><td>' + tag(u.total.judg) + '</td></tr>';
    });
    $('editor').innerHTML = h + '</tbody>';
  }
  function renderManual() {
    $('found').innerHTML = '<div><b>' + st.manual.length + '</b><span>rows to fill in</span></div>';
    $('headstrip').innerHTML = '<p class="muted small">Typing numbers in, so there are no columns to match.</p>'; $('colmap').innerHTML = '';
    var h = '<thead><tr><th>Grade</th><th>Phase</th><th>Students</th><th>Expected or better</th><th>Better than expected</th><th></th></tr></thead><tbody>';
    st.manual.forEach(function (r, i) {
      function inp(k, t) { return '<input type="' + (t || 'number') + '" data-man="' + i + '" data-k="' + k + '" value="' + (r[k] == null ? '' : esc(r[k])) + '" aria-label="' + esc(r.grade + ' ' + k) + '"' + (t ? ' style="width:70px"' : ' min="0" step="1" inputmode="numeric"') + '>'; }
      h += '<tr><td>' + inp('grade', 'text') + '</td><td>' + inp('phase', 'text') + '</td><td>' + inp('E') + '</td><td>' + inp('G') + '</td><td>' + inp('H') + '</td><td>' + (i === st.manual.length - 1 ? '<button class="btn quiet" type="button" id="add-man">Add Row</button>' : '') + '</td></tr>';
    });
    $('editor').innerHTML = h + '</tbody>';
  }
  $('editor').addEventListener('input', function (e) {
    var t = e.target;
    if (t.dataset.man != null) {
      var r = st.manual[+t.dataset.man], k = t.dataset.k;
      r[k] = (k === 'grade' || k === 'phase') ? t.value.trim() : (t.value === '' ? null : Math.max(0, Math.round(+t.value)));
      clearTimeout(timer); timer = setTimeout(function () { manualToSource(); renderStepper(); }, 200); return;
    }
    if (t.dataset.edit) {
      var id = t.dataset.edit, k2 = t.dataset.k, row = id.split(':');
      var v = t.value === '' ? null : Math.max(0, Math.round(+t.value));
      (st.edits[id] = st.edits[id] || {})[k2] = v;
      clearTimeout(timer); timer = setTimeout(function () { var y = window.scrollY; recompute(); window.scrollTo(0, y); }, 400);
    }
  });
  $('editor').addEventListener('click', function (e) {
    if (e.target.id === 'add-man') { var l = st.manual[st.manual.length - 1]; var m = l ? /(\d+)/.exec(l.grade) : null; st.manual.push({ grade: m ? 'G' + (+m[1] + 1) : 'G1', phase: l ? l.phase : 1, E: null, G: null, H: null }); renderManual(); }
  });

  /* ---------- step 3: rules ---------- */
  function setVal(id, v) { var el = $(id); if (el && document.activeElement !== el) el.value = v; }
  function renderRules() {
    var c = st.cfg;
    setVal('r-acc', Math.round(c.acc * 100)); setVal('r-vw', c.vw == null ? '' : Math.round(c.vw * 100));
    setVal('r-good', Math.round(c.good * 100)); setVal('r-vg', Math.round(c.vg * 100)); setVal('r-out', Math.round(c.out * 100));
    setVal('cap-w', Math.round(c.capWeak * 100)); setVal('cap-o', Math.round(c.capOut * 100)); setVal('subst', c.substantialPts); setVal('sim-title', c.simTitle);
    $('use-assessed').checked = c.useAssessed; $('vw-weak').checked = c.vwCountsAsWeak;
    document.querySelector('input[name="method"][value="' + c.totalMethod + '"]').checked = true;
    document.querySelector('input[name="round"][value="' + c.rounding + '"]').checked = true;
    document.querySelector('input[name="basis"][value="' + c.basis + '"]').checked = true;
    $('points-box').hidden = c.totalMethod !== 'points';
    $('round-flag').hidden = !(c.totalMethod === 'points' && c.rounding === 'unspecified');
    if (!$('pts').children.length) $('pts').innerHTML = E.CATS.slice().reverse().map(function (k) { return '<label class="pt">' + tag(k) + '<input type="number" min="0" max="20" step="1" data-pt="' + esc(k) + '" aria-label="Points for ' + esc(k) + '"></label>'; }).join('');
    Array.prototype.forEach.call(document.querySelectorAll('[data-pt]'), function (el) { if (document.activeElement !== el) el.value = c.points[el.dataset.pt] == null ? '' : c.points[el.dataset.pt]; });
    // ladders
    var z1 = [], acc = c.acc * 100, vw = c.vw == null ? null : c.vw * 100;
    if (vw != null && vw < acc) z1.push(['vw', vw, 'Very Weak'], ['weak', acc - vw, 'Weak']); else z1.push(['weak', acc, 'Weak']);
    z1.push(['acc', 100 - acc, 'Acceptable or better']);
    $('lad1').innerHTML = z1.map(function (z) { return '<span class="k-' + z[0] + '" style="flex-basis:' + Math.max(0, z[1]) + '%">' + (z[1] > 9 ? z[2] : '') + '</span>'; }).join('');
    var g = c.good * 100, v = c.vg * 100, o = c.out * 100;
    var z2 = [['acc', g, 'Acceptable'], ['good', v - g, 'Good'], ['vg', o - v, 'Very Good'], ['out', 100 - o, 'Outstanding']];
    $('lad2').innerHTML = z2.map(function (z) { return '<span class="k-' + z[0] + '" style="flex-basis:' + Math.max(0, z[1]) + '%">' + (z[1] > 7 ? z[2] : '') + '</span>'; }).join('');
    function scale(marks) { return [0].concat(marks).concat([100]).filter(function (x, i, a) { return x != null && a.indexOf(x) === i; }).map(function (x) { return '<span style="left:' + x + '%">' + Math.round(x) + '%</span>'; }).join(''); }
    $('lad1s').innerHTML = scale([vw, acc]); $('lad2s').innerHTML = scale([g, v, o]);
    // rounding example
    var avg = 2.5, ex = { unspecified: 'nearest (for now)', nearest: 'nearest', down: 'round down', up: 'round up' }[c.rounding];
    var rn = E.roundBy(avg, c.rounding);
    $('round-eg').innerHTML = 'Example: average 2.5<span class="arrow">→</span>' + ex + '<span class="arrow">→</span>' + rn + '<span class="arrow">→</span>' + tag(E.catFromPoints(rn, c));
    if (st.base) { ['both', 'grades', 'phases'].forEach(function (k) { $('n-' + k).textContent = E.counted(st.base.items, k).length; }); }
  }
  function num01(id, fb) { var v = parseFloat($(id).value); return isNaN(v) ? fb : Math.max(0, Math.min(100, v)) / 100; }
  ['r-acc', 'r-vw', 'r-good', 'r-vg', 'r-out', 'cap-w', 'cap-o'].forEach(function (id) {
    $(id).addEventListener('input', function () {
      var c = st.cfg, d = E.defaults();
      c.acc = num01('r-acc', d.acc); c.good = num01('r-good', d.good); c.vg = num01('r-vg', d.vg); c.out = num01('r-out', d.out);
      c.vw = $('r-vw').value === '' ? null : num01('r-vw', null);
      c.capWeak = num01('cap-w', d.capWeak); c.capOut = num01('cap-o', d.capOut);
      renderRules(); soon();
    });
  });
  $('subst').addEventListener('input', function () { var v = parseFloat(this.value); st.cfg.substantialPts = isNaN(v) ? 10 : v; soon(); });
  $('sim-title').addEventListener('input', function () { st.cfg.simTitle = this.value.trim() || E.defaults().simTitle; st.wb = null; });
  $('use-assessed').addEventListener('change', function () { st.cfg.useAssessed = this.checked; soon(); });
  $('vw-weak').addEventListener('change', function () { st.cfg.vwCountsAsWeak = this.checked; soon(); });
  $('pts').addEventListener('input', function (e) { var el = e.target.closest('[data-pt]'); if (!el) return; st.cfg.points[el.dataset.pt] = el.value === '' ? null : +el.value; renderRules(); soon(); });
  document.addEventListener('change', function (e) {
    var t = e.target;
    if (t.name === 'method') { st.cfg.totalMethod = t.value; renderRules(); soon(); }
    if (t.name === 'round') { st.cfg.rounding = t.value; renderRules(); soon(); }
    if (t.name === 'basis') { st.cfg.basis = t.value; st.sel = 0; renderRules(); soon(); }
  });
  $('reset-rules').addEventListener('click', function () { var cols = st.cfg.columns; st.cfg = E.defaults(); st.cfg.columns = cols; renderRules(); soon(); });

  /* ---------- shared visuals ---------- */
  function short(it) { return it.kind === 'total' ? 'P' + it.phase + ' total' : String(it.grade).replace(/^grade\s*/i, 'G'); }
  function tilesBlock(model, base, c) {
    var levels = [['Grades', 'grades'], ['Phase totals', 'phases']];
    var html = levels.map(function (lv) {
      var list = E.counted(model.items, lv[1]); if (!list.length) return '';
      var k = E.check(list, c);
      var baseIds = {}; if (base) base.items.forEach(function (x) { baseIds[x.id] = x; });
      var order = E.CATS; list = list.slice().sort(function (a, b) { return order.indexOf(a.judg) - order.indexOf(b.judg); });
      return '<div class="tiles-row"><div class="lab"><b>' + lv[0] + '</b><span class="muted">' + k.weak + ' of ' + k.n + ' Weak' + (c.basis === 'both' ? '' : (c.basis === lv[1] ? ', counted toward the limit' : ', not counted')) + '</span></div><ul class="tiles" aria-label="' + esc(lv[0]) + '">' +
        list.map(function (it) { var ch = base && baseIds[it.id] && baseIds[it.id].judg !== it.judg; return '<li class="tile" title="' + esc(it.label + ': ' + it.judg + (ch ? ' (was ' + baseIds[it.id].judg + ')' : '')) + '"><span class="sq k-' + (KEY[it.judg] || 'acc') + (ch ? ' changed' : '') + '"></span><span class="tl">' + esc(short(it)) + '</span></li>'; }).join('') + '</ul></div>';
    }).join('');
    var used = {}; model.items.forEach(function (x) { used[x.judg] = 1; });
    html += '<ul class="legend">' + E.CATS.slice().reverse().filter(function (k) { return used[k] || k !== 'Very Weak'; }).map(function (k) { return '<li>' + dot(k) + esc(k) + '</li>'; }).join('') + (base ? '<li><span class="sq changed" style="width:14px;height:14px;border-radius:3px;background:transparent"></span>Changed</li>' : '') + '</ul>';
    return html;
  }
  function viewer(gridId, tabsId, names, sel, which) {
    var wb = workbook(); if (!wb) return;
    $(tabsId).innerHTML = names.map(function (n, i) { return '<button type="button" role="tab" aria-selected="' + (i === sel) + '" data-tab="' + which + '" data-i="' + i + '">' + esc(n) + '</button>'; }).join('');
    var ws = wb.getWorksheet(names[sel]);
    $(gridId).innerHTML = ws ? Grid.render(ws) : '';
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-tab]'); if (!b) return;
    if (b.dataset.tab === 'a') st.tabA = +b.dataset.i; else st.tabS = +b.dataset.i;
    render();
  });

  /* ---------- step 4: results ---------- */
  function renderResults() {
    var c = st.cfg, m = st.base, k = E.check(E.counted(m.items, c.basis), c), fx = E.findFixes(st.parsed, c);
    var basisTxt = { both: 'grades and phase totals', grades: 'grades', phases: 'phase totals' }[c.basis];
    $('kpis').innerHTML =
      '<div class="kpi ' + (k.weakOk ? 'ok' : 'bad') + '"><span class="l">Weak</span><span class="v">' + k.weak + ' of ' + k.n + '</span><span class="s">' + pc(k.n ? k.weak / k.n : 0) + ' of ' + basisTxt + '. Limit: ' + k.maxW + '.</span></div>' +
      '<div class="kpi ' + (k.outOk ? 'ok' : 'bad') + '"><span class="l">Outstanding</span><span class="v">' + k.out + ' of ' + k.n + '</span><span class="s">' + pc(k.n ? k.out / k.n : 0) + '. Limit: ' + k.maxO + '.</span></div>' +
      '<div class="kpi"><span class="l">Fixed in the sheet</span><span class="v">' + fx.fixes.length + '</span><span class="s">' + (fx.fixes.length ? 'No student numbers changed' : 'Nothing needed fixing') + '</span></div>';
    $('tiles-actual').innerHTML = tilesBlock(m, null, c);
    // bars
    var h = '<div class="bars">';
    m.units.forEach(function (u) {
      if (m.units.length > 1) h += '<div class="muted small" style="margin-top:6px"><b>' + esc((st.parsed.blocks.length > 1 ? u.block + ', ' : '') + (u.group !== 'All students' ? u.group + ', ' : '') + 'Phase ' + u.phase) + '</b></div>';
      u.grades.forEach(function (g) {
        h += '<div class="bar-row"><span class="nm">' + esc(g.grade) + '</span><span class="bar-track"><span class="bar-fill k-' + (KEY[g.judg] || 'acc') + '" style="width:' + Math.max(0, Math.min(100, (g.I || 0) * 100)) + '%"></span><span class="bar-line" style="left:' + c.acc * 100 + '%"></span></span><span class="bar-val">' + pc(g.I) + ' ' + dot(g.judg) + '</span></div>';
      });
    });
    h += '</div><div class="bar-axis"><i></i><div><span style="left:0">0%</span><span style="left:' + c.acc * 100 + '%">' + Math.round(c.acc * 100) + '%</span><span style="left:100%">100%</span></div><i></i></div>';
    $('bars').innerHTML = h;
    // phase cards
    var cards = m.units.filter(function (u) { return u.total; }).map(function (u) {
      var t = u.total, title = (st.parsed.blocks.length > 1 ? u.block + ', ' : '') + (u.group !== 'All students' && m.units.length > 1 ? u.group + ', ' : '') + 'Phase ' + u.phase;
      var body;
      if (c.totalMethod === 'points') {
        body = '<div class="pts-row">' + u.grades.map(function (g) { return '<span class="chip" title="' + esc(g.grade + ': ' + g.judg) + '">' + dot(g.judg) + esc(short(g)) + ' ' + (g.pts == null ? '–' : g.pts) + '</span>'; }).join('') + '</div>' +
          '<div class="calc"><span>' + t.sumPts + ' ÷ ' + t.nPts + ' =</span><b>' + (t.avg == null ? '–' : t.avg.toFixed(2)) + '</b><span class="arrow">→</span><span>' + (t.rounded == null ? '' : t.rounded) + '</span><span class="arrow">→</span>' + tag(t.judg) + '</div>' +
          (t.roundingFlag ? '<div class="flag"><span class="dot warn"></span><span><strong>Confirm rounding.</strong> No rule given for decimals, so the nearest point is used for now.</span></div>' : '');
      } else body = '<div class="calc"><span>' + t.G + ' of ' + t.A + ' assessed</span><span class="arrow">→</span><b>' + pc(t.I) + '</b><span class="arrow">→</span>' + tag(t.judg) + '</div>';
      body += '<p class="muted small">Together: ' + t.G + ' of ' + t.A + ' assessed students made expected or better progress (' + pc(t.I) + '), ' + t.H + ' did better than expected (' + pc(t.J) + ').</p>';
      return '<div class="pc"><b>' + esc(title) + ' total</b>' + body + '</div>';
    });
    $('phase-card').hidden = !cards.length; $('phasecards').innerHTML = cards.join('');
    $('fix-card').hidden = !fx.fixes.length;
    $('fix-h').textContent = 'Fixed in the sheet (' + fx.fixes.length + ')';
    $('fixes').innerHTML = fx.fixes.map(function (f) { return '<li><span class="dot" aria-hidden="true"></span><span><b>' + esc(f.what) + '.</b> ' + esc(f.detail) + '</span></li>'; }).join('');
    viewer('xl-a-grid', 'xl-a-tabs', [Book.NAMES.orig, Book.NAMES.actual, Book.NAMES.log], Math.min(st.tabA, 2), 'a');
  }

  /* ---------- step 5: simulation ---------- */
  function renderSim() {
    var c = st.cfg, s = st.sim;
    if (!s.needed) { $('sim-status').innerHTML = '<div class="big-flag ok"><span class="dot ok"></span><span><b>Already within both limits.</b> The actual results meet the targets, so nothing needs simulating.</span></div>'; $('sim-body').hidden = true; return; }
    if (s.failed) { $('sim-status').innerHTML = '<div class="big-flag bad"><span class="dot bad"></span><span><b>No simulation can meet the limits with these rules.</b> ' + (c.basis === 'phases' ? 'With only phase totals counted, the limit may allow no Weak results at all. Try counting grades and totals together.' : 'Check the limits and which judgements count.') + '</span></div>'; $('sim-body').hidden = true; return; }
    $('sim-body').hidden = false;
    var o = s.options[st.sel];
    $('sim-status').innerHTML = '<div class="kpis"><div class="kpi sim"><span class="l">Judgements that change</span><span class="v">' + o.changed.length + '</span><span class="s">The fewest possible</span></div><div class="kpi sim"><span class="l">Students</span><span class="v">' + o.students + '</span><span class="s">' + pc(o.share) + ' of assessed students</span></div><div class="kpi ' + (o.substantial ? 'bad' : '') + '"><span class="l">Largest change in one grade</span><span class="v">' + o.maxPts.toFixed(1) + '</span><span class="s">percentage points' + (o.substantial ? ', substantial' : '') + '</span></div></div>';
    $('opts').innerHTML = s.options.map(function (x, i) { return '<button type="button" aria-pressed="' + (i === st.sel) + '" data-opt="' + i + '"><span class="o-n">Option ' + String.fromCharCode(65 + i) + '</span><span class="o-s">' + x.students + ' students, ' + x.changed.length + ' changes' + (i === 0 ? ', smallest' : '') + '</span></button>'; }).join('');
    $('opts-note').textContent = s.options.length > 1 ? 'All options change the fewest judgements possible. They differ in which grades move and how many students that takes.' : '';
    // before / after stacked bars
    function stack(list) {
      var d = E.dist(list), n = d.n || 1;
      return E.CATS.map(function (k) { var w = d.counts[k] / n * 100; return w ? '<span class="k-' + KEY[k] + '" style="flex-basis:' + w + '%" title="' + esc(k + ': ' + d.counts[k]) + '"></span>' : ''; }).join('');
    }
    var ba = '';
    [['Grades', 'grades'], ['Phase totals', 'phases'], ['All', 'both']].forEach(function (lv) {
      var la = E.counted(st.base.items, lv[1]), ls = E.counted(o.model.items, lv[1]); if (!la.length) return;
      var ka = E.check(la, c), ks = E.check(ls, c), capX = c.capWeak * 100;
      ba += '<div><div class="muted small" style="margin-bottom:4px"><b>' + lv[0] + '</b>' + (c.basis === lv[1] ? ' (counted toward the limit)' : '') + '</div>' +
        '<div class="ba-row"><span>Actual</span><span class="stack">' + stack(la) + '<i class="cap" style="left:' + capX + '%"></i></span><span class="' + (ka.weakOk ? '' : 'bad') + '" style="color:' + (ka.weakOk ? 'var(--ok)' : 'var(--bad)') + '">' + ka.weak + ' Weak</span></div>' +
        '<div class="ba-row" style="margin-top:6px"><span style="color:var(--sim);font-weight:600">Simulated</span><span class="stack">' + stack(ls) + '<i class="cap" style="left:' + capX + '%"></i></span><span style="color:' + (ks.weakOk ? 'var(--ok)' : 'var(--bad)') + '">' + ks.weak + ' Weak</span></div></div>';
    });
    $('ba').innerHTML = ba + '<p class="muted small">Red is Weak. The black line is the Weak limit (' + Math.round(c.capWeak * 100) + '%): red must stay left of it.</p>';
    // dumbbell of changed grades
    var vals = []; o.cells.forEach(function (x) { vals.push(x.from.I, x.to.I); }); vals.push(c.acc);
    var lo = Math.max(0, Math.floor((Math.min.apply(null, vals) * 100 - 5) / 10) * 10), hi = 100;
    function X(v) { return ((v * 100 - lo) / (hi - lo) * 100).toFixed(2); }
    $('dumb').innerHTML = o.cells.map(function (x) {
      var a = x.from.I, b = x.to.I;
      return '<div class="db-row"><span class="nm" style="font-weight:560">' + esc(x.to.grade) + '</span><span class="db-track"><span class="db-line" style="left:' + X(c.acc) + '%"></span><span class="db-seg" style="left:' + X(Math.min(a, b)) + '%;width:' + (Math.abs(b - a) * 100 / (hi - lo) * 100).toFixed(2) + '%"></span><span class="db-a" style="left:' + X(a) + '%"></span><span class="db-b" style="left:' + X(b) + '%"></span></span><span class="db-v">' + pc(a) + ' → <b>' + pc(b) + '</b><br><span class="muted small">+' + x.dG + ' of ' + x.from.A + '</span></span></div>';
    }).join('') + '<div class="db-row muted small"><span></span><span class="db-track" style="height:16px"><span style="position:absolute;left:0">' + lo + '%</span><span style="position:absolute;left:' + X(c.acc) + '%;transform:translateX(-50%)">' + Math.round(c.acc * 100) + '% line</span><span style="position:absolute;right:0">100%</span></span><span></span></div>';
    // assumptions
    var items = o.cells.map(function (x) {
      return '<li><span class="dot sim"></span><span><b>' + esc(x.to.grade) + ':</b> assume ' + (x.dG ? x.dG + ' more of its ' + x.from.A + ' assessed students make expected progress (' + x.from.G + ' to ' + x.to.G + ')' : '') + (x.dH ? (x.dG ? ' and ' : '') + Math.abs(x.dH) + (x.dH > 0 ? ' more' : ' fewer') + ' do better than expected' : '') + '. ' + tag(x.from.judg) + ' <span class="arrow">→</span> ' + tag(x.to.judg) + (x.pp > c.substantialPts ? ' <b style="color:var(--bad)">Substantial: ' + x.pp.toFixed(1) + ' points</b>' : '') + '</span></li>';
    });
    o.changed.filter(function (x) { return x.from.kind === 'total'; }).forEach(function (x) { items.push('<li><span class="dot"></span><span><b>' + esc(x.to.label) + '</b> changes on its own because its grades changed. ' + tag(x.from.judg) + ' <span class="arrow">→</span> ' + tag(x.to.judg) + '</span></li>'); });
    items.push('<li><span class="dot"></span><span>Everything else keeps its actual value. Cohort sizes, assessed totals and thresholds are unchanged.</span></li>');
    $('assume').innerHTML = items.join('');
    $('tiles-sim').innerHTML = tilesBlock(o.model, st.base, c);
    viewer('xl-s-grid', 'xl-s-tabs', [Book.simSheetName(c), Book.NAMES.comp, Book.NAMES.counts, Book.NAMES.totals], Math.min(st.tabS, 3), 's');
  }
  $('opts').addEventListener('click', function (e) { var b = e.target.closest('[data-opt]'); if (!b) return; st.sel = +b.dataset.opt; st.wb = null; render(); });

  /* ---------- step 6: download ---------- */
  function renderDownload() {
    var c = st.cfg, opt = st.sim.options && st.sim.options[st.sel];
    var list = [
      ['#808080', Book.NAMES.orig, 'Your sheet, exactly as loaded'],
      ['#548235', Book.NAMES.actual, 'Same layout, mistakes fixed in orange, formulas throughout'],
      opt ? ['#c00000', Book.simSheetName(c), 'Same layout, simulated counts in blue, labelled on every page'] : null,
      ['#2e75b6', Book.NAMES.comp, 'Every count: actual, simulated, difference, judgement change'],
      ['#2e75b6', Book.NAMES.counts, 'Counts and percentages for grades, totals and both, checked against both limits'],
      ['#2e75b6', Book.NAMES.totals, 'Points, average, rounding and total judgement for each phase'],
      ['#548235', Book.NAMES.log, 'Every correction made to the actual results'],
      opt ? ['#c00000', Book.NAMES.plan, 'Each simulated change written as a planning assumption, with its size'] : null,
      ['#2e75b6', Book.NAMES.settings, 'The rules used. Change them in Excel and the formulas follow']
    ].filter(Boolean);
    $('wb-list').innerHTML = list.map(function (x) { return '<li style="--c:' + x[0] + '"><b>' + esc(x[1]) + '</b><span>' + esc(x[2]) + '</span></li>'; }).join('');
    var flags = [];
    if (st.base.items.some(function (x) { return x.roundingFlag; }) || (opt && opt.model.items.some(function (x) { return x.roundingFlag; }))) flags.push('Some phase averages have decimals and no rounding rule was given. They are marked in the workbook for someone to confirm.');
    if (opt && opt.substantial) flags.push('The simulation needs a substantial change (' + opt.maxPts.toFixed(1) + ' percentage points in one grade). This is stated in the workbook.');
    $('dl-flag').hidden = !flags.length; $('dl-flag-t').textContent = flags.join(' ');
  }

  /* ---------- downloads ---------- */
  var dlP = null;
  function downloads() {
    var plain = { save: function (req) { return new Promise(function (res) { var url = URL.createObjectURL(new Blob([req.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); var a = document.createElement('a'); a.href = url; a.download = req.filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(url); }, 4000); res({ status: 'saved' }); }); } };
    if (!dlP) dlP = (window.claude && typeof window.claude.use === 'function') ? window.claude.use('downloads').catch(function () { return null; }) : Promise.resolve(plain);
    return dlP;
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-dl]'); if (!b) return;
    var kind = b.dataset.dl, out = kind === 'template' ? 'msg1' : 'msg6';
    if (typeof ExcelJS === 'undefined') { msg(out, 'The Excel writer did not load. Reload the page and try again.', 'bad'); return; }
    var wb, name;
    try {
      if (kind === 'template') { wb = Book.template(); name = 'Progress Judgements - Blank Template.xlsx'; }
      else { wb = workbook(); var subj = (st.parsed.blocks[0] && st.parsed.blocks[0].title) || 'Progress'; name = (subj + ' - Actual and Simulated Results.xlsx').replace(/[\\/:*?"<>|]/g, '-'); }
    } catch (err) { msg(out, 'The workbook could not be built (' + err.message + ').', 'bad'); return; }
    msg(out, 'Preparing ' + name + '…');
    Promise.all([wb.xlsx.writeBuffer(), downloads()]).then(function (r) {
      if (!r[1]) { msg(out, 'Downloads are not available in this view.', 'bad'); return; }
      return r[1].save({ filename: name, data: new Uint8Array(r[0]) }).then(function () { msg(out, 'Saved ' + name + '.', 'ok'); }, function (err) { var c = err && err.code; msg(out, c === 'declined' ? 'Download cancelled.' : 'The download did not go through. Try again.', c === 'declined' ? '' : 'bad'); });
    }).catch(function (err) { msg(out, 'The workbook could not be built (' + (err && err.message) + ').', 'bad'); });
  });
  $('start-over').addEventListener('click', function () { st.src = null; st.origWb = null; st.headerRow = null; st.grid = null; st.parsed = null; st.base = null; st.sim = null; st.manual = null; st.edits = {}; st.wb = null; $('title').textContent = 'Progress Judgements'; $('sheet-pick').hidden = true; msg('msg1', 'Cleared. Nothing from that sheet is kept.', 'ok'); go(1); });

/*ASSISTANT*/
  go(1);
})();
