(function () {
  var E = Engine;
  var $ = function (id) { return document.getElementById(id); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  var KEY = { 'Very Weak': 'vw', 'Weak': 'weak', 'Acceptable': 'acc', 'Good': 'good', 'Very Good': 'vg', 'Outstanding': 'out' };
  function dot(j) { return '<span class="dot k-' + (KEY[j] || 'acc') + '" aria-hidden="true"></span>'; }
  function tag(j) { return '<span class="judg">' + dot(j) + esc(j) + '</span>'; }
  function pc(v, d) { if (d == null) d = st && (st.cfg.pctDecimals === 0 || st.cfg.pctDecimals === 2) ? st.cfg.pctDecimals : 1; return v == null ? '–' : (v * 100).toFixed(d) + '%'; }

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
    st.rawFile = st.pendingRaw || null; st.pendingRaw = null;
    st.src = wb; st.origWb = orig || null; st.headerRow = null; st.srcName = name; st.edits = {}; st.manual = null; st.sel = 0; st.wb = null; st.cfg.simLocked = []; st.cfg.simPinned = {};
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
    st.tabsWithTables = withTables.map(function (ws) { return ws.name; });
    if (withTables.length > 1) {
      $('sheet-pick').hidden = false;
      $('sheet-list').innerHTML = '<button type="button" class="btn" data-sheet="' + ALL_TABS + '">All ' + withTables.length + ' Sheets Together</button>' + withTables.map(function (ws, i) { return '<button type="button" class="btn' + (i === 0 ? ' primary' : '') + '" data-sheet="' + esc(ws.name) + '">' + esc(ws.name) + '</button>'; }).join('');
    } else $('sheet-pick').hidden = true;
    pickSheet(withTables[0].name);
    msg('msg1', withTables.length > 1 ? 'Loaded ' + name + '. It has ' + withTables.length + ' sheets with results; showing “' + withTables[0].name + '”. Choose another, or all sheets together so the limits apply across every sheet.' : 'Loaded ' + name + '.', 'ok');
    go(2);
  }
  // Several tabs (for example one per subject) stacked into one sheet, so the limits apply across all of them.
  var ALL_TABS = '__all__';
  function combineSheets(wb, names) {
    var ws = wb.getWorksheet('All sheets') || wb.addWorksheet('All sheets'), r0 = 0;
    names.forEach(function (n) {
      var src = wb.getWorksheet(n); if (!src || src === ws) return;
      var t = ws.getCell(r0 + 1, 1); t.value = n; t.font = { bold: true, size: 13 }; r0 += 2;
      var maxR = 0;
      src.eachRow({ includeEmpty: false }, function (row, r) {
        maxR = Math.max(maxR, r);
        if (row.height) ws.getRow(r0 + r).height = row.height;
        row.eachCell({ includeEmpty: false }, function (cell, c) {
          var v = cell.value; if (cell.isMerged && cell.master !== cell) v = null;
          if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) v = v.result == null ? null : v.result;
          var d = ws.getCell(r0 + r, c); d.value = v; try { d.style = JSON.parse(JSON.stringify(cell.style || {})); } catch (e) { }
        });
      });
      (src.model.merges || []).forEach(function (m) {
        var mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m); if (!mm) return;
        try { ws.mergeCells(mm[1] + (+mm[2] + r0) + ':' + mm[3] + (+mm[4] + r0)); } catch (e) { }
      });
      for (var c = 1; c <= (src.columnCount || 0); c++) { var w = src.getColumn(c).width; if (w && (!ws.getColumn(c).width || ws.getColumn(c).width < w)) ws.getColumn(c).width = w; }
      r0 += maxR + 2;
    });
    return ws.name;
  }
  function useAllSheets() {
    var names = (st.tabsWithTables || []).filter(function (n) { return n !== 'All sheets'; });
    if (names.length < 2) return false;
    var n1 = combineSheets(st.src, names); if (st.origWb) combineSheets(st.origWb, names);
    pickSheet(n1); msg('msg1', 'Using all ' + names.length + ' sheets together: ' + names.join(', ') + '.', 'ok');
    Array.prototype.forEach.call(document.querySelectorAll('[data-sheet]'), function (b) { b.classList.toggle('primary', b.dataset.sheet === ALL_TABS); });
    return true;
  }
  function pickSheet(name) {
    st.sheetName = name; st.headerRow = null; st.grid = gridFromWs(st.src.getWorksheet(name)); st.cfg.columns = {}; st.edits = {}; st.sel = 0; st.cfg.simLocked = []; st.cfg.simPinned = {};
    Array.prototype.forEach.call(document.querySelectorAll('[data-sheet]'), function (b) { b.classList.toggle('primary', b.dataset.sheet === name); });
    recompute();
  }
  var SHEET_EXT = /^(xlsx|xlsm|xls|xlsb|ods|fods|csv|tsv|txt|numbers)$/;
  // Loads spreadsheet bytes into the page. Resolves {ok, error}.
  function loadSheetBuffer(buf, name, ext) {
    ext = (ext || (name.split('.').pop() || '')).toLowerCase();
    st.pendingRaw = { buf: buf.slice(0), name: name, ext: ext, sent: false }; // the file as uploaded, for the assistant's sandbox
    return new Promise(function (resolve) {
      function fail(t) { st.pendingRaw = null; msg('msg1', t, 'bad'); resolve({ ok: false, error: t }); }
      if (buf.byteLength < 200 && (ext === 'xlsx' || ext === 'xlsm')) return fail('This file is almost empty (' + buf.byteLength + ' bytes), so it is not a real spreadsheet. It was probably a failed download. Download it again and check it opens in Excel.');
      if (ext === 'xlsx' || ext === 'xlsm') {
        if (typeof ExcelJS === 'undefined') return fail('The spreadsheet reader did not load. Reload the page, or enter the rows instead.');
        var wb = new ExcelJS.Workbook(), orig = new ExcelJS.Workbook();
        Promise.all([wb.xlsx.load(buf.slice(0)), orig.xlsx.load(buf.slice(0))]).then(function () { useWorkbook(wb, name, orig); resolve({ ok: !!st.src }); }, function () { fail('This does not look like a working Excel file. Open it in Excel, save a fresh copy, and try again.'); });
        return;
      }
      try {
        if (typeof XLSX === 'undefined') throw new Error('reader not loaded');
        var x = /^(csv|tsv|txt)$/.test(ext) ? XLSX.read(new TextDecoder().decode(new Uint8Array(buf)), { type: 'string' }) : XLSX.read(new Uint8Array(buf), { type: 'array' });
        var sheets = x.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(x.Sheets[n], { header: 1, raw: true, defval: null }) }; });
        useWorkbook(aoaToWorkbook(sheets), name, aoaToWorkbook(sheets));
        resolve({ ok: !!st.src });
      } catch (err) { fail(ext === 'numbers' ? 'That Numbers file could not be read. In Numbers, choose File, Export To, Excel, and use that file.' : 'That file could not be read. Try saving it as .xlsx, or enter the rows.'); }
    });
  }
  function readFile(f) {
    if (!f) return;
    var ext = (f.name.split('.').pop() || '').toLowerCase();
    if (!SHEET_EXT.test(ext)) { msg('msg1', 'That is not a spreadsheet. For a PDF, Word file or picture of a table, attach it to the assistant instead and it will read it.', 'bad'); return; }
    f.arrayBuffer().then(function (buf) { loadSheetBuffer(buf, f.name, ext); }, function () { msg('msg1', 'That file could not be read. If it came from a download, download it again and check it opens in Excel.', 'bad'); });
  }
  // Google Sheets: a link shared as "Anyone with the link" can be read.
  function gsheetId(text) { var m = /docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]{20,})/.exec(String(text || '')); return m ? m[1] : null; }
  function loadGoogleSheet(link) {
    var id = gsheetId(link);
    if (!id) return Promise.resolve({ ok: false, error: 'That is not a Google Sheets link.' });
    var viaSite = !(window.claude && typeof window.claude.use === 'function');
    var url = viaSite ? '/api/gsheet?id=' + encodeURIComponent(id) : 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
    return fetch(url).then(function (r) {
      if (!r.ok) throw new Error(r.status === 404 || r.status === 403 || r.status === 401 ? 'private' : 'http ' + r.status);
      return r.arrayBuffer();
    }).then(function (buf) {
      var head = new Uint8Array(buf.slice(0, 4));
      if (!(head[0] === 0x50 && head[1] === 0x4b)) throw new Error('private');
      return loadSheetBuffer(buf, 'Google Sheet.xlsx', 'xlsx');
    }).catch(function (e) {
      var t = e && e.message === 'private'
        ? 'That Google Sheet is not shared publicly, so it cannot be opened from here. In Google Sheets choose Share, General access, Anyone with the link. Or choose File, Download, Microsoft Excel and load that file.'
        : 'The Google Sheet could not be opened from here. In Google Sheets choose File, Download, Microsoft Excel (.xlsx) and load that file instead.';
      msg('msg1', t, 'bad'); return { ok: false, error: t };
    });
  }
  $('gs-go').addEventListener('click', function () { var v = $('gs-link').value.trim(); if (!v) { $('gs-link').focus(); return; } loadGoogleSheet(v); });
  $('gs-link').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); $('gs-go').click(); } });
  $('file-in').addEventListener('change', function (e) { readFile(e.target.files && e.target.files[0]); e.target.value = ''; });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { readFile(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]); });
  $('sheet-list').addEventListener('click', function (e) { var b = e.target.closest('[data-sheet]'); if (!b) return; if (b.dataset.sheet === ALL_TABS) useAllSheets(); else pickSheet(b.dataset.sheet); });
  /* ---------- entering rows by hand (or pasting cells from Excel) ---------- */
  var ROWS = { kind: 'counts', title: '', list: [] };
  var RKEYS = ['phase', 'grade', 'E', 'G', 'H'];
  function blankRow() { return { phase: '', grade: '', E: '', G: '', H: '' }; }
  function ensureRows(n) { while (ROWS.list.length < n) ROWS.list.push(blankRow()); }
  function rowsHead() { return ROWS.kind === 'pct' ? ['Phase', 'Grade', 'Students', '% expected or better', '% better than expected'] : ['Phase', 'Grade', 'Students', 'Expected or better', 'Better than expected']; }
  var X_ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>';
  function renderRows() {
    ensureRows(4);
    var hd = rowsHead();
    var h = '<thead><tr><th></th>' + hd.map(function (t) { return '<th scope="col">' + esc(t) + '</th>'; }).join('') + '<th><span class="sr">Remove</span></th></tr></thead><tbody>';
    ROWS.list.forEach(function (r, i) {
      h += '<tr><td class="n">' + (i + 1) + '</td>' + RKEYS.map(function (k, j) {
        var numK = j >= 2;
        return '<td><input type="text" data-ri="' + i + '" data-rk="' + k + '"' + (numK ? ' inputmode="decimal"' : '') + ' value="' + esc(r[k]) + '" aria-label="Row ' + (i + 1) + ' ' + esc(hd[j]) + '" autocomplete="off"' + (k === 'grade' && i === 0 ? ' placeholder="Grade 1"' : k === 'phase' && i === 0 ? ' placeholder="2"' : '') + '></td>';
      }).join('') + '<td class="x"><button type="button" class="rm" data-rm="' + i + '" aria-label="Remove row ' + (i + 1) + '">' + X_ICON + '</button></td></tr>';
    });
    $('rows-tbl').innerHTML = h + '</tbody>';
    $('rows-title').value = ROWS.title;
    Array.prototype.forEach.call(document.querySelectorAll('input[name="rows-kind"]'), function (x) { x.checked = x.value === ROWS.kind; });
  }
  function openRows(open) {
    var c = $('rows-card'); c.hidden = !open; $('open-rows').setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) { renderRows(); c.scrollIntoView({ block: 'start', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' }); var f = c.querySelector('input[data-rk="grade"]'); if (f) f.focus({ preventScroll: true }); }
  }
  $('open-rows').addEventListener('click', function () { openRows($('rows-card').hidden); });
  $('rows-add').addEventListener('click', function () { ROWS.list.push(blankRow()); renderRows(); var ins = $('rows-tbl').querySelectorAll('input[data-rk="grade"]'); ins[ins.length - 1].focus(); });
  $('rows-clear').addEventListener('click', function () { ROWS.list = []; ROWS.title = ''; renderRows(); });
  $('rows-title').addEventListener('input', function () { ROWS.title = this.value; });
  Array.prototype.forEach.call(document.querySelectorAll('input[name="rows-kind"]'), function (x) { x.addEventListener('change', function () { ROWS.kind = this.value; renderRows(); }); });
  $('rows-tbl').addEventListener('input', function (e) { var t = e.target; if (t.dataset.ri == null) return; ROWS.list[+t.dataset.ri][t.dataset.rk] = t.value; });
  $('rows-tbl').addEventListener('click', function (e) { var b = e.target.closest('[data-rm]'); if (!b) return; ROWS.list.splice(+b.dataset.rm, 1); renderRows(); });
  function cellVal(v) { v = String(v == null ? '' : v).trim(); return v; }
  $('rows-tbl').addEventListener('paste', function (e) {
    var t = e.target; if (t.dataset.ri == null) return;
    var txt = (e.clipboardData || window.clipboardData).getData('text'); if (!txt || !/[\t\n]/.test(txt.replace(/\n$/, ''))) return;
    e.preventDefault();
    var lines = txt.replace(/\r/g, '').replace(/\n+$/, '').split('\n');
    var aoa = lines.map(function (l) { return l.split('\t').map(function (v) { v = v.trim(); return v === '' ? null : (isNaN(+v) ? v : +v); }); });
    // A pasted block that includes its own headings goes through the full reader
    if (E.parse(E.gridFromAoa(aoa), {}).blocks.length) { var w1 = aoaToWorkbook([{ name: 'Pasted', aoa: aoa }]); openRows(false); useWorkbook(w1, 'pasted cells', aoaToWorkbook([{ name: 'Pasted', aoa: aoa }])); return; }
    var r0 = +t.dataset.ri, c0 = RKEYS.indexOf(t.dataset.rk);
    lines.forEach(function (l, i) {
      ensureRows(r0 + i + 1);
      l.split('\t').forEach(function (v, j) { var k = RKEYS[c0 + j]; if (k) ROWS.list[r0 + i][k] = cellVal(v); });
    });
    renderRows();
    msg('msg1', 'Pasted ' + lines.length + ' row' + (lines.length === 1 ? '' : 's') + '.', 'ok');
  });
  function numIn(v) { v = String(v == null ? '' : v).replace(/[,\s%]/g, ''); if (v === '') return null; var n = +v; return isFinite(n) ? n : NaN; }
  function useRows(src) {
    var pctK = ROWS.kind === 'pct', out = [], bad = [];
    ROWS.list.forEach(function (r, i) {
      if (!String(r.grade || '').trim() && !String(r.E || '').trim()) return;
      var E1 = numIn(r.E), G1 = numIn(r.G), H1 = numIn(r.H), n = i + 1;
      if (!String(r.grade || '').trim()) { bad.push('row ' + n + ' has no grade'); return; }
      if (!(E1 > 0) || Math.round(E1) !== E1) { bad.push('row ' + n + ' needs a whole number of students'); return; }
      if (G1 == null || isNaN(G1) || H1 == null || isNaN(H1)) { bad.push('row ' + n + ' is missing a progress number'); return; }
      if (pctK) {
        if (G1 > 100 || H1 > 100) { bad.push('row ' + n + ' has a percentage over 100'); return; }
        if (H1 > G1) { bad.push('row ' + n + ': better than expected cannot be more than expected or better'); return; }
        out.push({ phase: String(r.phase || '').trim(), grade: String(r.grade).trim(), E: E1, I: G1, J: H1 });
      } else {
        if (G1 > E1 || H1 > E1) { bad.push('row ' + n + ' has more students making progress than in the group'); return; }
        if (H1 > G1) { bad.push('row ' + n + ': better than expected cannot be more than expected or better'); return; }
        out.push({ phase: String(r.phase || '').trim(), grade: String(r.grade).trim(), E: E1, G: G1, H: H1 });
      }
    });
    if (bad.length) { msg('msg1', 'Please check: ' + bad.join('; ') + '.', 'bad'); return { ok: false, problems: bad }; }
    if (!out.length) { msg('msg1', 'Fill in at least one row: a grade, the number of students and the two progress numbers.', 'bad'); return { ok: false, problems: ['no rows'] }; }
    var title = ROWS.title.trim() || 'Progress';
    var w = Book.synthetic(out, title, { pct: pctK }), o = Book.synthetic(out, title, { pct: pctK });
    useWorkbook(w, src || 'rows typed in', o);
    return { ok: true, rows: out.length };
  }
  $('rows-use').addEventListener('click', function () { useRows(); });

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
    var subj = subjectName();
    $('title').textContent = subj && subj !== 'Progress' ? (/progress$/i.test(subj) ? subj + ' Judgements' : subj + ' Progress Judgements') : 'Progress Judgements';
    render();
  }
  function workbook() {
    // Rebuilt for each download so the summary and assistant notes are current.
    if (st.base) st.wb = Book.build({ srcWb: st.src, origWb: st.origWb, sheetName: st.sheetName, parsed: st.parsed, cfg: st.cfg, base: st.base, sim: st.sim, sel: st.sel, summary: summaryOf(), notes: notesOf() });
    return st.wb;
  }
  function subjectName() {
    var b = st.parsed && st.parsed.blocks; if (!b || !b.length) return 'Progress';
    if (b.length === 1) return b[0].title || st.sheetName || 'Progress';
    // Several tables: name them by what their titles share, or by the sheet.
    var t = b.map(function (x) { return String(x.title || ''); }), pre = t[0];
    t.forEach(function (x) { while (pre && x.indexOf(pre) !== 0) pre = pre.slice(0, -1); });
    pre = pre.replace(/[\s:,\-–]+$/, '').trim();
    if (pre.length >= 4) return pre;
    // Generic tab names ("Sheet1") say nothing, so use the file name instead.
    if (st.sheetName && !/^(sheet|tab|data|feuil|tabelle|hoja)\s*\d*$|^all sheets$/i.test(st.sheetName)) return st.sheetName;
    var fn = String(st.srcName || '').replace(/\.[a-z0-9]+$/i, '').replace(/^[0-9a-f]{8}-/i, '').replace(/[_]+/g, ' ').replace(/\s+\d{1,3}$/, '').replace(/\s*\(\d+\)$/, '').trim();
    return fn || st.sheetName || 'Progress';
  }
  function fileBase() { return subjectName().replace(/[\\/:*?"<>|]/g, '-'); }
  // One plain-language account of the results, used by the PDF report and the workbook's Summary sheet.
  function summaryOf() {
    var c = st.cfg, m = st.base, k = E.check(E.counted(m.items, c.basis), c), fx = E.findFixes(st.parsed, c).fixes;
    var basisTxt = { both: 'grade and phase total judgements', grades: 'grade judgements', phases: 'phase total judgements' }[c.basis];
    var P = function (v) { return Math.round(v * 1000) / 10 + '%'; };
    var secs = [];
    var verdict = k.weakOk && k.outOk && k.extraOk ? 'The actual results meet the targets.' : 'The actual results do not meet the targets.';
    secs.push({ heading: 'Result', text: verdict + ' ' + k.weak + ' of ' + k.n + ' ' + basisTxt + ' are Weak (' + P(k.n ? k.weak / k.n : 0) + '; at most ' + k.maxW + ' allowed). ' + k.out + ' are Outstanding (' + P(k.n ? k.out / k.n : 0) + '; at most ' + k.maxO + ' allowed).' + (k.extra.length ? ' ' + k.extra.map(function (x) { return x.label + ': ' + (x.ok ? 'met' : 'not met') + ' (' + x.count + ' of ' + k.n + ').'; }).join(' ') : '') });
    var d = k.d, cats = E.CATS.slice().reverse();
    secs.push({ heading: 'How many of each judgement (actual)', table: { head: ['Judgement', 'Count', 'Share'], rows: cats.map(function (x) { return [x, d.counts[x], P(d.n ? d.counts[x] / d.n : 0)]; }).concat([['Total', d.n, '100%']]) } });
    secs.push({ heading: 'Rules used', bullets: [
      'Acceptable: at least ' + P(c.acc) + ' make expected or better progress' + (c.vw != null ? '; Very Weak below ' + P(c.vw) : '') + '.',
      'Good, Very Good, Outstanding: at least ' + P(c.good) + ', ' + P(c.vg) + ', ' + P(c.out) + ' make better than expected progress.',
      'Limits: Weak at most ' + P(c.capWeak) + ', Outstanding at most ' + P(c.capOut) + ', counted over ' + basisTxt + '.',
      'Phase totals: ' + (c.totalMethod === 'points' ? 'average of grade points, rounding ' + (c.rounding === 'unspecified' ? 'not given (nearest point used for now, to be confirmed)' : c.rounding) : 'from summed student counts') + '.'
    ].concat((c.extraRules || []).map(function (r) { return 'Extra rule: ' + (r.label || E.ruleLabel(r)) + '.'; })) });
    secs.push({ heading: 'Corrections to the sheet', bullets: fx.length ? fx.map(function (f) { return f.what + ': ' + f.detail; }) : ['Nothing needed correcting.'] });
    var totals = m.items.filter(function (x) { return x.kind === 'total'; });
    if (totals.length) secs.push({ heading: 'Phase totals', table: { head: ['Phase', 'Average points', 'Rounded', 'Judgement', 'Note'], rows: totals.map(function (t) { return [(t.block && st.parsed.blocks.length > 1 ? t.block + ': ' : '') + t.label, t.avg == null ? '' : Math.round(t.avg * 100) / 100, t.rounded == null ? '' : t.rounded, t.judg, t.roundingFlag ? 'Rounding to confirm' : '']; }) } });
    secs.push({ heading: 'Every judgement (actual)', table: { head: ['Grade or total', 'Students', 'Assessed', 'Expected or better', 'Better than expected', 'Judgement'], rows: m.items.map(function (it) { return [(st.parsed.blocks.length > 1 ? it.block + ': ' : '') + (it.kind === 'total' ? it.label : it.grade), it.E, it.A, it.G + ' (' + P(it.I || 0) + ')', it.H + ' (' + P(it.J || 0) + ')', it.judg]; }) } });
    var s = st.sim, o = s && s.options && s.options[st.sel];
    if (!s || !s.needed) secs.push({ heading: 'Simulated scenario', text: 'Not needed: the actual results already meet the targets.' });
    else if (s.failed) secs.push({ heading: 'Simulated scenario', text: 'No simulated scenario can meet the targets with these rules and limits.', tone: 'sim' });
    else {
      var ks = o.check;
      secs.push({ heading: c.simTitle, tone: 'sim', text: 'For internal planning only. These are planning assumptions, not actual results. Option ' + String.fromCharCode(65 + st.sel) + ' of ' + s.options.length + ': ' + o.changed.length + ' judgements change, ' + o.students + ' students (' + P(o.share) + ' of assessed students). Largest change in one grade: ' + o.maxPts.toFixed(1) + ' percentage points' + (o.substantial ? ', which is a substantial adjustment' : '') + '. With these assumptions, ' + ks.weak + ' of ' + ks.n + ' are Weak and ' + ks.out + ' Outstanding.' });
      secs.push({ heading: 'Planning assumptions', tone: 'sim', table: { head: ['Grade', 'Assumed change', 'From', 'To', 'Points', 'Judgement'], rows: o.cells.map(function (x) { return [(st.parsed.blocks.length > 1 ? x.to.block + ': ' : '') + x.to.grade, (x.dG ? (x.dG > 0 ? '+' : '') + x.dG + ' expected or better' : '') + (x.dH ? (x.dG ? ', ' : '') + (x.dH > 0 ? '+' : '') + x.dH + ' better than expected' : ''), x.from.G + ' / ' + x.from.H, x.to.G + ' / ' + x.to.H, x.pp.toFixed(1), x.from.judg + ' to ' + x.to.judg]; }) } });
      var tc = o.changed.filter(function (x) { return x.from.kind === 'total'; });
      if (tc.length) secs.push({ heading: 'Phase totals that change as a result', tone: 'sim', bullets: tc.map(function (x) { return x.to.label + ': ' + x.from.judg + ' to ' + x.to.judg + '.'; }) });
    }
    var src = st.srcName ? 'Source: ' + st.srcName + (st.sheetName ? ', sheet "' + st.sheetName + '"' : '') + '. ' : '';
    return { title: subjectName() + ': progress judgements', subtitle: src + 'Prepared ' + new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) + '.', sections: secs };
  }
  function notesOf() {
    var turns = (typeof AS !== 'undefined' && AS.turns) ? AS.turns : [], log = (typeof AS !== 'undefined' && AS.log) ? AS.log : [];
    return { transcript: turns.map(function (t) { return { role: t.role, text: String(t.content).replace(/\n<<<BEGIN [\s\S]*?END [^\n]*>>>/g, '\n[attached document text]') }; }), changes: log.slice().reverse().map(function (x) { return x.label; }) };
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
        h += '<tr><td>' + esc(g.grade) + '</td><td><input type="text" data-edit="' + esc(g.id) + '" data-k="phase" value="' + esc(g.phase) + '" aria-label="' + esc(g.grade + ' phase') + '" style="width:64px"></td><td>' + inp('E') + '</td><td>' + (g.A == null ? '' : g.A) + '</td><td>' + inp('G') + '</td><td>' + inp('H') + '</td><td>' + tag(g.judg) + '</td></tr>';
      });
      if (u.total) h += '<tr class="tot"><td>Total</td><td>' + esc(u.phase) + '</td><td>' + u.total.E + '</td><td>' + u.total.A + '</td><td>' + u.total.G + '</td><td>' + u.total.H + '</td><td>' + tag(u.total.judg) + '</td></tr>';
    });
    $('editor').innerHTML = h + '</tbody>';
  }
  $('editor').addEventListener('input', function (e) {
    var t = e.target;
    if (t.dataset.edit) {
      var id = t.dataset.edit, k2 = t.dataset.k, row = id.split(':');
      var v = k2 === 'phase' ? (t.value.trim() || null) : (t.value === '' ? null : Math.max(0, Math.round(+t.value)));
      (st.edits[id] = st.edits[id] || {})[k2] = v;
      clearTimeout(timer); timer = setTimeout(function () { var y = window.scrollY; recompute(); window.scrollTo(0, y); }, 400);
    }
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
    var dpIn = document.querySelector('input[name="pctdp"][value="' + (c.pctDecimals == null ? 1 : c.pctDecimals) + '"]'); if (dpIn) dpIn.checked = true;
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
    var xk = st.base ? E.check(E.counted(st.base.items, c.basis), c).extra : [];
    $('xrules').innerHTML = (c.extraRules || []).map(function (r, i) {
      var k = xk[i];
      return '<li><span class="dot ' + (k ? (k.ok ? 'ok' : 'bad') : '') + '" aria-hidden="true"></span><span>' + esc(r.label || E.ruleLabel(r)) + '</span><span class="st">' + (k ? (k.ok ? 'Met' : 'Not met') + ': ' + k.count + ' of ' + E.counted(st.base.items, c.basis).length + (r.op === 'min' ? ', need ' + k.limit : ', allowed ' + k.limit) : '') + '</span><button type="button" class="btn quiet" data-xr-rm="' + i + '">Remove</button></li>';
    }).join('') || '<li style="background:transparent;padding:0;display:block" class="muted">None yet.</li>';
  }
  $('xr-add').addEventListener('click', function () {
    var v = parseFloat($('xr-share').value); if (isNaN(v) || v < 0 || v > 100) { $('xr-share').focus(); return; }
    st.cfg.extraRules = (st.cfg.extraRules || []).concat([{ op: $('xr-op').value, share: v / 100, cats: $('xr-cats').value.split(',') }]);
    $('xr-share').value = ''; st.sel = 0; renderRules(); soon();
  });
  $('xrules').addEventListener('click', function (e) { var b = e.target.closest('[data-xr-rm]'); if (!b) return; st.cfg.extraRules.splice(+b.dataset.xrRm, 1); renderRules(); soon(); });
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
    if (t.name === 'pctdp') { st.cfg.pctDecimals = +t.value; st.wb = null; renderRules(); soon(); }
  });
  $('reset-rules').addEventListener('click', function () { var cols = st.cfg.columns; st.cfg = E.defaults(); st.cfg.columns = cols; renderRules(); soon(); });

  /* ---------- shared visuals ---------- */
  function short(it) { return it.kind === 'total' ? 'P' + it.phase + ' total' : String(it.grade).replace(/^grade\s*/i, 'G'); }
  function barsBlock(m, c) {
    var h = '<div class="bars">';
    m.units.forEach(function (u) {
      if (m.units.length > 1) h += '<div class="muted small" style="margin-top:6px"><b>' + esc((st.parsed.blocks.length > 1 ? u.block + ', ' : '') + (u.group !== 'All students' ? u.group + ', ' : '') + E.phaseLabel(u.phase)) + '</b></div>';
      u.grades.forEach(function (g) {
        h += '<div class="bar-row"><span class="nm">' + esc(g.grade) + '</span><span class="bar-track"><span class="bar-fill k-' + (KEY[g.judg] || 'acc') + '" style="width:' + Math.max(0, Math.min(100, (g.I || 0) * 100)) + '%"></span><span class="bar-line" style="left:' + c.acc * 100 + '%"></span></span><span class="bar-val">' + pc(g.I) + ' ' + dot(g.judg) + '</span></div>';
      });
    });
    h += '</div><div class="bar-axis"><i></i><div><span style="left:0">0%</span><span style="left:' + c.acc * 100 + '%">' + Math.round(c.acc * 100) + '%</span><span style="left:100%">100%</span></div><i></i></div>';
    return h;
  }
  function phaseCards(m, c) {
    return m.units.filter(function (u) { return u.total; }).map(function (u) {
      var t = u.total, title = (st.parsed.blocks.length > 1 ? u.block + ', ' : '') + (u.group !== 'All students' && m.units.length > 1 ? u.group + ', ' : '') + E.phaseLabel(u.phase);
      var body;
      if (c.totalMethod === 'points') {
        body = '<div class="pts-row">' + u.grades.map(function (g) { return '<span class="chip" title="' + esc(g.grade + ': ' + g.judg) + '">' + dot(g.judg) + esc(short(g)) + ' ' + (g.pts == null ? '–' : g.pts) + '</span>'; }).join('') + '</div>' +
          '<div class="calc"><span>' + t.sumPts + ' ÷ ' + t.nPts + ' =</span><b>' + (t.avg == null ? '–' : t.avg.toFixed(2)) + '</b><span class="arrow">→</span><span>' + (t.rounded == null ? '' : t.rounded) + '</span><span class="arrow">→</span>' + tag(t.judg) + '</div>' +
          (t.roundingFlag ? '<div class="flag"><span class="dot warn"></span><span><strong>Confirm rounding.</strong> No rule given for decimals, so the nearest point is used for now.</span></div>' : '');
      } else body = '<div class="calc"><span>' + t.G + ' of ' + t.A + ' assessed</span><span class="arrow">→</span><b>' + pc(t.I) + '</b><span class="arrow">→</span>' + tag(t.judg) + '</div>';
      body += '<p class="muted small">Together: ' + t.G + ' of ' + t.A + ' assessed students made expected or better progress (' + pc(t.I) + '), ' + t.H + ' did better than expected (' + pc(t.J) + ').</p>';
      return '<div class="pc"><b>' + esc(title) + ' total</b>' + body + '</div>';
    });
  }
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
      k.extra.map(function (x) { return '<div class="kpi ' + (x.ok ? 'ok' : 'bad') + '"><span class="l">' + esc(x.label) + '</span><span class="v">' + x.count + ' of ' + k.n + '</span><span class="s">' + (x.op === 'min' ? 'Needs ' : 'Allowed: ') + x.limit + '. ' + (x.ok ? 'Met.' : 'Not met.') + '</span></div>'; }).join('') +
      '<div class="kpi"><span class="l">Fixed in the sheet</span><span class="v">' + fx.fixes.length + '</span><span class="s">' + (fx.fixes.length ? 'No student numbers changed' : 'Nothing needed fixing') + '</span></div>';
    $('tiles-actual').innerHTML = tilesBlock(m, null, c);
    $('bars').innerHTML = barsBlock(m, c);
    var cards = phaseCards(m, c);
    $('phase-card').hidden = !cards.length; $('phasecards').innerHTML = cards.join('');
    $('fix-card').hidden = !fx.fixes.length;
    $('fix-h').textContent = 'Fixed in the sheet (' + fx.fixes.length + ')';
    $('fixes').innerHTML = fx.fixes.map(function (f) { return '<li><span class="dot" aria-hidden="true"></span><span><b>' + esc(f.what) + '.</b> ' + esc(f.detail) + '</span></li>'; }).join('');
    viewer('xl-a-grid', 'xl-a-tabs', [Book.NAMES.orig, Book.NAMES.actual, Book.NAMES.log], Math.min(st.tabA, 2), 'a');
  }

  /* ---------- step 5: simulation ---------- */
  function renderSimLimits() {
    var c = st.cfg, grades = st.base.items.filter(function (x) { return x.kind === 'grade'; }), multi = st.parsed.blocks.length > 1;
    if (document.activeElement !== $('sl-stu')) $('sl-stu').value = c.simMaxStudents == null ? '' : c.simMaxStudents;
    if (document.activeElement !== $('sl-pts')) $('sl-pts').value = c.simMaxPts == null ? '' : c.simMaxPts;
    $('sl-lock').innerHTML = grades.map(function (g) { return '<label><input type="checkbox" class="vh" data-lock="' + esc(g.id) + '"' + ((c.simLocked || []).indexOf(g.id) >= 0 ? ' checked' : '') + '>' + esc((multi ? g.block + ': ' : '') + g.grade) + '</label>'; }).join('');
    var byId = {}; grades.forEach(function (g) { byId[g.id] = g; });
    var pins = Object.keys(c.simPinned || {}).filter(function (k) { return byId[k]; });
    $('sl-pinned').innerHTML = pins.length ? '<div class="field" style="margin-top:var(--s3)"><span>Your own simulated changes</span><ul class="xrules">' + pins.map(function (k) { var g = byId[k], p = c.simPinned[k]; return '<li><span class="dot sim" aria-hidden="true"></span><span><b>' + esc(g.grade) + '</b>' + (p.G != null ? ': expected or better ' + g.G + ' → ' + p.G : '') + (p.H != null ? (p.G != null ? ',' : ':') + ' better than expected ' + g.H + ' → ' + p.H : '') + '</span><span></span><button type="button" class="btn quiet" data-unpin="' + esc(k) + '">Remove</button></li>'; }).join('') + '</ul></div>' : '';
    var parts = [];
    if (c.simMaxStudents != null) parts.push('at most ' + c.simMaxStudents + ' students per grade');
    if (c.simMaxPts != null) parts.push('at most ' + c.simMaxPts + ' points per grade');
    if ((c.simLocked || []).length) parts.push(c.simLocked.length + ' grade' + (c.simLocked.length === 1 ? '' : 's') + ' left alone');
    if (pins.length) parts.push(pins.length + ' change' + (pins.length === 1 ? '' : 's') + ' of your own');
    $('simlim-sum').textContent = parts.length ? parts.join(', ') : 'None set. Any grade can change.';
    if (parts.length) $('simlim').open = true;
  }
  function numOrNull(v) { var n = parseFloat(v); return isNaN(n) || n < 0 ? null : n; }
  $('sl-stu').addEventListener('input', function () { st.cfg.simMaxStudents = numOrNull(this.value); st.sel = 0; soon(); });
  $('sl-pts').addEventListener('input', function () { st.cfg.simMaxPts = numOrNull(this.value); st.sel = 0; soon(); });
  $('sl-lock').addEventListener('change', function (e) { var t = e.target; if (!t.dataset.lock) return; var L = (st.cfg.simLocked || []).filter(function (x) { return x !== t.dataset.lock; }); if (t.checked) L.push(t.dataset.lock); st.cfg.simLocked = L; st.sel = 0; recompute(); });
  $('sl-pinned').addEventListener('click', function (e) { var b = e.target.closest('[data-unpin]'); if (!b) return; delete st.cfg.simPinned[b.dataset.unpin]; st.sel = 0; recompute(); });
  function renderSim() {
    var c = st.cfg, s = st.sim;
    renderSimLimits();
    if (!s.needed) { $('sim-status').innerHTML = '<div class="big-flag ok"><span class="dot ok"></span><span><b>Already within both limits.</b> The actual results meet the targets, so nothing needs simulating.</span></div>'; $('sim-body').hidden = true; return; }
    if (s.failed && s.limited) { $('sim-status').innerHTML = '<div class="big-flag bad"><span class="dot bad"></span><span><b>No option fits inside your limits for the simulation.</b> Loosen them below, for example allow more students per grade or let another grade change.</span></div>'; $('sim-body').hidden = true; return; }
    if (s.failed) { $('sim-status').innerHTML = '<div class="big-flag bad"><span class="dot bad"></span><span><b>No simulation can meet the limits with these rules.</b> ' + (c.basis === 'phases' ? 'With only phase totals counted, the limit may allow no Weak results at all. Try counting grades and totals together.' : 'Check the limits and which judgements count.') + '</span></div>'; $('sim-body').hidden = true; return; }
    $('sim-body').hidden = false;
    var o = s.options[st.sel];
    $('sim-status').innerHTML = '<div class="kpis"><div class="kpi sim"><span class="l">Judgements that change</span><span class="v">' + o.changed.length + '</span><span class="s">The fewest possible</span></div><div class="kpi sim"><span class="l">Students</span><span class="v">' + o.students + '</span><span class="s">' + pc(o.share) + ' of assessed students</span></div><div class="kpi ' + (o.substantial ? 'bad' : '') + '"><span class="l">Largest change in one grade</span><span class="v">' + o.maxPts.toFixed(1) + '</span><span class="s">percentage points' + (o.substantial ? ', substantial' : '') + '</span></div></div>';
    $('opts').innerHTML = s.options.map(function (x, i) { return '<button type="button" aria-pressed="' + (i === st.sel) + '" data-opt="' + i + '"><span class="o-n">Option ' + String.fromCharCode(65 + i) + '</span><span class="o-s">' + x.students + ' students, ' + x.changed.length + ' changes' + (i === 0 ? ', smallest' : '') + '</span></button>'; }).join('');
    var notes = [];
    if (s.pinned) notes.push('Includes your own simulated changes.');
    if (s.options.length > 1) notes.push('All options change the fewest judgements possible. They differ in which grades move and how many students that takes.');
    var xf = o.check.extra.filter(function (x) { return !x.ok; });
    if (xf.length) notes.push('Not met in this option: ' + xf.map(function (x) { return x.label; }).join('; ') + '. The simulation only fixes Weak and Outstanding.');
    $('opts-note').textContent = notes.join(' ');
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
      ['#2e75b6', Book.NAMES.settings, 'The rules used. Change them in Excel and the formulas follow'],
      ['#262626', Book.NAMES.summary, 'The results in plain words, first in the workbook'],
      (typeof AS !== 'undefined' && (AS.turns.length || AS.log.length)) ? ['#0000da', Book.NAMES.notes, 'What the assistant was asked, said and changed'] : null
    ].filter(Boolean);
    $('wb-list').innerHTML = list.map(function (x) { return '<li style="--c:' + x[0] + '"><b>' + esc(x[1]) + '</b><span>' + esc(x[2]) + '</span></li>'; }).join('');
    var flags = [];
    if (st.base.items.some(function (x) { return x.roundingFlag; }) || (opt && opt.model.items.some(function (x) { return x.roundingFlag; }))) flags.push('Some phase averages have decimals and no rounding rule was given. They are marked in the workbook for someone to confirm.');
    if (opt && opt.substantial) flags.push('The simulation needs a substantial change (' + opt.maxPts.toFixed(1) + ' percentage points in one grade). This is stated in the workbook.');
    $('dl-flag').hidden = !flags.length; $('dl-flag-t').textContent = flags.join(' ');
    var hasChat = typeof AS !== 'undefined' && (AS.turns.length || AS.log.length);
    $('dl-chat').hidden = !hasChat;
  }

/*REPORT*/
  /* ---------- downloads ---------- */
  var dlP = null;
  var MIME = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pdf: 'application/pdf', md: 'text/markdown', txt: 'text/plain' };
  function downloads() {
    var plain = { save: function (req) { return new Promise(function (res) { var ext = req.filename.split('.').pop(); var url = URL.createObjectURL(req.data instanceof Blob ? req.data : new Blob([req.data], { type: MIME[ext] || 'application/octet-stream' })); var a = document.createElement('a'); a.href = url; a.download = req.filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(url); }, 4000); res({ status: 'saved' }); }); } };
    if (!dlP) dlP = (window.claude && typeof window.claude.use === 'function') ? window.claude.use('downloads').catch(function () { return null; }) : Promise.resolve(plain);
    return dlP;
  }
  function saveFile(name, dataP, out) {
    msg(out, 'Preparing ' + name + '…');
    return Promise.all([dataP, downloads()]).then(function (r) {
      if (!r[1]) { msg(out, 'Downloads are not available in this view.', 'bad'); return false; }
      return r[1].save({ filename: name, data: r[0] }).then(function () { msg(out, 'Saved ' + name + '.', 'ok'); return true; }, function (err) { var c = err && err.code; msg(out, c === 'declined' ? 'Download cancelled.' : 'The download did not go through. Try again.', c === 'declined' ? '' : 'bad'); return false; });
    }).catch(function (err) { msg(out, name + ' could not be made (' + (err && err.message || err) + ').', 'bad'); return false; });
  }
  var PDF_LIBS = ['https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js'];
  var pdfLibP = null;
  function loadPdfLibs() {
    if (!pdfLibP) pdfLibP = PDF_LIBS.reduce(function (p, url) { return p.then(function () { return new Promise(function (res, rej) { var el = document.createElement('script'); el.src = url; el.onload = res; el.onerror = function () { pdfLibP = null; rej(new Error('the PDF writer did not load')); }; document.head.appendChild(el); }); }); }, Promise.resolve());
    return pdfLibP;
  }
  // A4 PDF of the same summary as the workbook's Summary sheet, plus the assistant's notes.
  function reportPdf() {
    var sm = summaryOf(), nt = notesOf();
    return loadPdfLibs().then(function () {
      var doc = new window.jspdf.jsPDF({ unit: 'pt', format: 'a4' }), W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight(), M = 48, y = M;
      var INK = [38, 38, 38], MUTED = [110, 110, 115], SIM = [0, 0, 218];
      function need(h) { if (y + h > H - M) { doc.addPage(); y = M; } }
      function para(t, size, color, bold) { doc.setFont('helvetica', bold ? 'bold' : 'normal'); doc.setFontSize(size); doc.setTextColor.apply(doc, color); var lines = doc.splitTextToSize(String(t), W - 2 * M); lines.forEach(function (l) { need(size * 1.35); doc.text(l, M, y + size); y += size * 1.35; }); }
      para(sm.title, 18, INK, true); y += 2; para(sm.subtitle, 9.5, MUTED); y += 10;
      sm.sections.forEach(function (sec) {
        need(40); y += 6; para(sec.heading, 12.5, sec.tone === 'sim' ? SIM : INK, true); y += 3;
        if (sec.text) { para(sec.text, 10, INK); y += 4; }
        (sec.bullets || []).forEach(function (b) { var lines = doc.splitTextToSize(String(b), W - 2 * M - 12); doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor.apply(doc, INK); lines.forEach(function (l, i) { need(13.5); if (i === 0) doc.text('\u2022', M, y + 10); doc.text(l, M + 12, y + 10); y += 13.5; }); });
        if (sec.table) {
          doc.autoTable({ startY: y + 2, head: [sec.table.head], body: sec.table.rows.map(function (r) { return r.map(function (v) { return v == null ? '' : String(v); }); }), margin: { left: M, right: M }, theme: 'grid',
            styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 4, lineColor: [225, 225, 228], lineWidth: 0.5, textColor: INK },
            headStyles: { fillColor: sec.tone === 'sim' ? [231, 231, 252] : [244, 244, 245], textColor: INK, fontStyle: 'bold' },
            didParseCell: function (h) { if (h.section === 'body' && /^(Weak|Very Weak)$/.test(h.cell.raw)) h.cell.styles.textColor = [192, 0, 0]; } });
          y = doc.lastAutoTable.finalY + 8;
        }
      });
      if (nt.changes.length || nt.transcript.length) {
        doc.addPage(); y = M; para('Assistant notes', 14, INK, true); para('What the assistant was asked, what it said, and what it changed on the page. The results above already include these changes.', 9.5, MUTED); y += 8;
        if (nt.changes.length) { para('Changes made', 11, INK, true); nt.changes.forEach(function (c, i) { para((i + 1) + '. ' + c, 10, INK); }); y += 8; }
        nt.transcript.forEach(function (m) { need(30); para(m.role === 'user' ? 'Asked' : 'Assistant', 10, m.role === 'user' ? SIM : INK, true); para(String(m.text).replace(/\*\*/g, '').slice(0, 12000), 9.5, INK); y += 6; });
      }
      var n = doc.getNumberOfPages();
      for (var i = 1; i <= n; i++) { doc.setPage(i); doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor.apply(doc, MUTED); doc.text(sm.title + '   Page ' + i + ' of ' + n, M, H - 24); }
      return doc.output('arraybuffer');
    });
  }
  function chatText() {
    var nt = notesOf(), lines = ['# ' + subjectName() + ': assistant conversation', '', 'Saved ' + new Date().toLocaleString('en-GB') + '.', ''];
    if (nt.changes.length) { lines.push('## Changes made', ''); nt.changes.forEach(function (c, i) { lines.push((i + 1) + '. ' + c); }); lines.push(''); }
    lines.push('## Conversation', '');
    nt.transcript.forEach(function (m) { lines.push('**' + (m.role === 'user' ? 'Asked' : 'Assistant') + ':**', '', m.text, ''); });
    return lines.join('\n');
  }
  // Exposed so the assistant can offer download buttons in its replies.
  function download(kind, out) {
    out = out || 'msg6';
    if (kind === 'template') { if (typeof ExcelJS === 'undefined') { msg('msg1', 'The Excel writer did not load. Reload the page and try again.', 'bad'); return; } return saveFile('Progress Judgements - Blank Template.xlsx', Book.template().xlsx.writeBuffer(), 'msg1'); }
    if (!st.base) { msg(out, 'Load a sheet first.', 'bad'); return; }
    if (kind === 'workbook') { if (typeof ExcelJS === 'undefined') { msg(out, 'The Excel writer did not load. Reload the page and try again.', 'bad'); return; } var wb; try { wb = workbook(); } catch (err) { msg(out, 'The workbook could not be built (' + err.message + ').', 'bad'); return; } return saveFile(fileBase() + ' - Actual and Simulated Results.xlsx', wb.xlsx.writeBuffer().then(function (b) { return new Uint8Array(b); }), out); }
    if (kind === 'report') return reportDownload(out);
    if (kind === 'chat') return saveFile(fileBase() + ' - Assistant conversation.md', Promise.resolve(chatText()), out);
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-dl]'); if (!b) return;
    download(b.dataset.dl, b.dataset.out || (b.dataset.dl === 'template' ? 'msg1' : 'msg6'));
  });
  $('start-over').addEventListener('click', function () { st.src = null; st.origWb = null; st.headerRow = null; st.grid = null; st.parsed = null; st.base = null; st.sim = null; st.manual = null; st.edits = {}; st.wb = null; $('title').textContent = 'Progress Judgements'; $('sheet-pick').hidden = true; msg('msg1', 'Cleared. Nothing from that sheet is kept.', 'ok'); go(1); });

/*ASSISTANT*/
  go(1);
})();
