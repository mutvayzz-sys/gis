(function () {
  var C = Core, V = View;
  function clone(x) { return JSON.parse(JSON.stringify(x)); }
  // Everything lives in this object, in memory only. Nothing is stored or sent.
  var state = { parsed: null, rows: [], rules: clone(C.DEFAULT_RULES), sel: 0, wb1: 0, edited: false };
  var last = null;
  var $ = function (id) { return document.getElementById(id); };
  function setHTML(id, html) { var el = $(id); if (el && el.innerHTML !== html) el.innerHTML = html; }
  function msg(text, kind) { var m = $('msg'); m.textContent = text || ''; m.className = 'msg' + (kind ? ' ' + kind : ''); }
  function hasRaw() { return !!(state.parsed && state.parsed.raw && state.parsed.raw.length); }

  function render(opts) {
    opts = opts || {};
    var v = V.all(state.parsed, state.rows, state.rules, state.sel, state.wb1);
    last = v;
    if (opts.editor) setHTML('editor-body', v.EDITOR); else refreshErrors();
    setHTML('n-all', v.N_ALL); setHTML('n-grades', v.N_GRADES); setHTML('n-totals', v.N_TOTALS);
    document.body.classList.toggle('is-empty', !!v.empty);
    var subj = state.parsed && state.parsed.subject;
    $('page-title').textContent = subj ? subj + ' Progress Judgements' : 'Progress Judgements';
    if (v.empty) return;
    state.sel = v.sel;
    setHTML('stats', v.STATS);
    $('source').textContent = state.parsed ? state.parsed.source : 'Entered by hand';
    setHTML('wb1-grid', hasRaw() ? v.WB1_GRID : C.renderGrid(C.sheetSpecModel(v.model, state.rules, null, subj || 'Your sheet'), { label: 'Sheet' }));
    setHTML('wb1-tabs', hasRaw() ? v.WB1_TABS : '');
    setHTML('wb1-notes', hasRaw() ? v.WB1_NOTES : '');
    $('wb1-grid').setAttribute('aria-labelledby', 'wtab-' + state.wb1);
    setHTML('sheet-none', v.SHEET_NONE);
    $('wb2').hidden = !v.WB2_GRID;
    setHTML('wb2-grid', v.WB2_GRID); setHTML('wb2-tabs', v.WB2_TABS); setHTML('wb2-changes', v.WB2_CHANGES);
    $('wb2-grid').setAttribute('aria-labelledby', 'otab-' + state.sel);
    setHTML('dist-now', v.DIST_NOW);
    setHTML('table-now', v.TABLE_NOW);
    var fixes = state.parsed ? v.FIXES : '<p class="muted small">These numbers were entered by hand, so there is no file to check.</p>';
    if (state.edited && state.parsed) fixes = '<p class="muted small">You have changed numbers since loading. This list checks the sheet as it was loaded.</p>' + fixes;
    setHTML('fixes', fixes);
    setHTML('compare', v.COMPARE);
    setHTML('sim-head', v.SIM_HEAD);
    setHTML('sim-tabs', v.SIM_TABS);
    $('opt-panel').hidden = !v.OPT_BODY;
    $('opt-panel').setAttribute('aria-labelledby', 'tab-' + state.sel);
    setHTML('opt-title', v.OPT_TITLE);
    setHTML('opt-body', v.OPT_BODY);
  }
  // Update row messages without rebuilding inputs, so focus stays put while typing.
  function refreshErrors() {
    var errs = V.validate(state.rows), body = $('editor-body');
    Array.prototype.slice.call(body.querySelectorAll('tr.err')).forEach(function (tr) { tr.parentNode.removeChild(tr); });
    Array.prototype.slice.call(body.querySelectorAll('tr')).forEach(function (tr, i) {
      var bad = errs[i] && errs[i].length && errs[i][0] !== 'empty';
      Array.prototype.slice.call(tr.querySelectorAll('input[type="number"]')).forEach(function (inp) {
        if (inp.dataset.f === 'phase') return;
        if (bad) inp.setAttribute('aria-invalid', 'true'); else inp.removeAttribute('aria-invalid');
      });
      if (bad) {
        var er = document.createElement('tr'); er.className = 'err';
        er.innerHTML = '<td></td><td colspan="5" class="err-msg">' + C.esc(errs[i].join(' ')) + ' This row is left out until fixed.</td>';
        tr.parentNode.insertBefore(er, tr.nextSibling);
      }
    });
  }
  var timer = null;
  function soon() { clearTimeout(timer); timer = setTimeout(function () { render(); }, 120); }

  /* ---------- loading ---------- */
  function loadParsed(p, label) {
    if (!p.rows.length) { msg('No grade rows were found. The sheet needs a Grade column and the three student counts.', 'bad'); return; }
    state.parsed = p; state.rows = p.rows.map(function (r) { return { grade: r.grade, phase: r.phase, E: r.E, G: r.G, H: r.H }; });
    state.edited = false; state.sel = 0; state.wb1 = 0;
    render({ editor: true });
    msg('Loaded ' + p.rows.length + ' grades from ' + label + '.', 'ok');
  }
  function readFile(f) {
    if (!f) return;
    var isText = /\.(csv|tsv|txt)$/i.test(f.name);
    var rd = new FileReader();
    rd.onerror = function () { msg('That file couldn’t be read. Try saving it again, or paste the rows instead.', 'bad'); };
    rd.onload = function () {
      try {
        if (isText) { loadParsed(C.parseGrid(C.parseText(rd.result), { source: f.name }), f.name); return; }
        if (typeof XLSX === 'undefined') { msg('Excel files can’t be opened right now. Copy the table in Excel and use Paste From Excel instead.', 'bad'); return; }
        var wb = XLSX.read(new Uint8Array(rd.result), { type: 'array', cellFormula: true });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        var hasF = Object.keys(ws).some(function (k) { return k[0] !== '!' && ws[k] && ws[k].f; });
        loadParsed(C.parseGrid(aoa, { source: f.name, hasFormulas: hasF }), f.name);
      } catch (err) { msg('That file couldn’t be read (' + err.message + '). Try pasting the rows instead.', 'bad'); }
    };
    if (isText) rd.readAsText(f); else rd.readAsArrayBuffer(f);
  }
  $('file-in').addEventListener('change', function (e) { readFile(e.target.files && e.target.files[0]); e.target.value = ''; });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { readFile(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]); });
  $('use-paste').addEventListener('click', function () {
    var t = $('paste-box').value;
    if (!t.trim()) { msg('Paste some rows first.', 'bad'); $('paste-box').focus(); return; }
    loadParsed(C.parseGrid(C.parseText(t), { source: 'Pasted data' }), 'the pasted rows');
    $('paste-box').value = '';
  });
  $('blank').addEventListener('click', function () {
    state.parsed = null; state.edited = true; state.sel = 0;
    state.rows = []; for (var g = 1; g <= 12; g++) state.rows.push({ grade: 'G' + g, phase: g <= 5 ? 2 : g <= 8 ? 3 : 4, E: null, G: null, H: null });
    render({ editor: true });
    msg('Blank rows are ready in Edit Your Data, further down. Fill in the numbers and the sheets appear here.', 'ok');
    var inp = $('e-E-0'); if (inp) inp.focus();
  });
  $('clear').addEventListener('click', function () {
    state.parsed = null; state.rows = []; state.edited = false; state.sel = 0; state.wb1 = 0;
    render({ editor: true }); msg('Cleared. Nothing from that sheet is kept.', 'ok');
  });

  /* ---------- editing ---------- */
  $('editor-body').addEventListener('input', function (e) {
    var t = e.target; if (!t.dataset || t.dataset.i == null) return;
    var row = state.rows[+t.dataset.i], f = t.dataset.f;
    if (f === 'grade') row.grade = t.value.trim();
    else { var n = t.value === '' ? null : Number(t.value); row[f] = (n == null || isNaN(n)) ? null : n; }
    state.edited = true; soon();
  });
  $('editor-body').addEventListener('click', function (e) {
    var b = e.target.closest('[data-rm]'); if (!b) return;
    var i = +b.dataset.rm, name = C.gradeName(state.rows[i].grade || 'row');
    state.rows.splice(i, 1); state.edited = true;
    render({ editor: true }); msg('Removed ' + name + '.', 'ok');
  });
  $('add-row').addEventListener('click', function () {
    var lastRow = state.rows[state.rows.length - 1];
    var m = lastRow ? /([0-9]+)/.exec(lastRow.grade) : null;
    state.rows.push({ grade: m ? 'G' + (+m[1] + 1) : 'G1', phase: lastRow ? lastRow.phase : 1, E: null, G: null, H: null });
    state.edited = true; render({ editor: true });
    var inp = $('e-E-' + (state.rows.length - 1)); if (inp) inp.focus();
  });

  /* ---------- rules ---------- */
  function ruleVal(id, fb) { var v = parseFloat($(id).value); return isNaN(v) ? fb : Math.max(0, Math.min(100, v)) / 100; }
  ['r-acc', 'r-good', 'r-vg', 'r-out', 'cap-weak', 'cap-out'].forEach(function (id) {
    $(id).addEventListener('input', function () {
      var r = state.rules, d = C.DEFAULT_RULES;
      r.acc = ruleVal('r-acc', d.acc); r.good = ruleVal('r-good', d.good); r.vg = ruleVal('r-vg', d.vg); r.out = ruleVal('r-out', d.out);
      r.capWeak = ruleVal('cap-weak', d.capWeak); r.capOut = ruleVal('cap-out', d.capOut);
      var ae = $('acc-echo'); if (ae) ae.textContent = String(Math.round(r.acc * 1000) / 10);
      state.sel = 0; soon();
    });
  });
  $('mode-ctl').addEventListener('change', function (e) { if (e.target.name === 'mode') { state.rules.mode = e.target.value; state.sel = 0; render(); } });

  /* ---------- tabs ---------- */
  function selectTab(prefix, i) {
    if (prefix === 'wtab') state.wb1 = i; else state.sel = i;
    render();
    var nb = $(prefix + '-' + i); if (nb) nb.focus();
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[role="tab"]'); if (!b) return;
    var list = b.closest('[role="tablist"]'); if (!list) return;
    selectTab(list.dataset.prefix || 'tab', +b.dataset.opt);
  });
  document.addEventListener('keydown', function (e) {
    var b = e.target.closest && e.target.closest('[role="tab"]'); if (!b) return;
    var list = b.closest('[role="tablist"]'); if (!list) return;
    var n = list.querySelectorAll('[role="tab"]').length, i = +b.dataset.opt;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') i = (i + 1) % n;
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') i = (i - 1 + n) % n;
    else if (e.key === 'Home') i = 0; else if (e.key === 'End') i = n - 1; else return;
    e.preventDefault(); selectTab(list.dataset.prefix || 'tab', i);
  });

  /* ---------- downloads (built here, in the browser) ---------- */
  var dlP = null;
  function getDownloads() {
    var plain = { save: function (req) { return new Promise(function (res) {
      var url = URL.createObjectURL(new Blob([req.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      var a = document.createElement('a'); a.href = url; a.download = req.filename; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000); res({ status: 'saved' }); }); } };
    if (!dlP) dlP = (window.claude && typeof window.claude.use === 'function') ? window.claude.use('downloads').catch(function () { return null; }) : Promise.resolve(plain);
    return dlP;
  }
  function dlMsg(kind, text, cls) { var el = $(kind === 'one' || kind === 'all' ? 'dl-msg2' : (kind === 'template' ? 'msg' : 'dl-msg1')); el.textContent = text; el.className = 'msg' + (cls ? ' ' + cls : ''); }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-dl]'); if (!b) return;
    var kind = b.dataset.dl;
    if (typeof ExcelJS === 'undefined') { dlMsg(kind, 'The Excel writer did not load. Reload the page and try again.', 'bad'); return; }
    if (kind !== 'template' && (!last || last.empty)) { dlMsg(kind, 'Load a sheet first.', 'bad'); return; }
    var subj = (state.parsed && state.parsed.subject) || 'Progress';
    var names = { template: 'Progress Judgements - Blank Template.xlsx', fixed: subj + ' Progress - Fixed.xlsx', one: subj + ' Progress - Option ' + C.letter(state.sel) + ' (simulated).xlsx', all: subj + ' Progress - Full Workbook.xlsx' };
    var name = names[kind].replace(/[\\/:*?"<>|]/g, '-');
    dlMsg(kind, 'Preparing ' + name + '…');
    var wb;
    try { wb = XLWriter.build(kind, { parsed: state.parsed, model: last && last.model, sim: last && last.sim, rules: state.rules, sel: state.sel }); }
    catch (err) { dlMsg(kind, 'The workbook couldn’t be built (' + err.message + ').', 'bad'); return; }
    wb.calcProperties = { fullCalcOnLoad: true };
    Promise.all([wb.xlsx.writeBuffer(), getDownloads()]).then(function (res) {
      var dl = res[1];
      if (!dl) { dlMsg(kind, 'Downloads are not available in this view.', 'bad'); return; }
      return dl.save({ filename: name, data: new Uint8Array(res[0]) }).then(function () { dlMsg(kind, 'Saved ' + name + '.', 'ok'); }, function (err) {
        var c = err && err.code;
        dlMsg(kind, c === 'declined' ? 'Download cancelled.' : c === 'rate_limited' ? 'Another download is waiting for you to confirm it.' : 'The download didn’t go through. Try again in a moment.', c === 'declined' ? '' : 'bad');
      });
    }).catch(function (err) { dlMsg(kind, 'The workbook couldn’t be built (' + (err && err.message) + ').', 'bad'); });
  });
  getDownloads().then(function (dl) { if (!dl) { document.body.classList.add('dl-off'); $('dl-note').hidden = false; } });

  render({ editor: true });
})();
