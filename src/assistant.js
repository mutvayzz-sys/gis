  /* ---------- the assistant (runs inside the app scope) ---------- */
  var AS = { turns: [], log: [], busy: false, ctl: null, mode: null };
  var openAssistantWith = null;
  var INSTRUCTIONS = window.__ASSISTANT_PROMPT__ || '';
  function snapshot() { return JSON.parse(JSON.stringify({ cfg: st.cfg, edits: st.edits, sheetName: st.sheetName, headerRow: st.headerRow || null, sel: st.sel })); }
  function restore(s) { st.cfg = s.cfg; st.edits = s.edits; st.headerRow = s.headerRow; st.sel = s.sel; if (s.sheetName !== st.sheetName && st.src.getWorksheet(s.sheetName)) { st.sheetName = s.sheetName; } st.grid = gridFromWs(st.src.getWorksheet(st.sheetName)); recompute(); }
  function record(label, before, extra) { AS.log.unshift({ id: Date.now() + Math.random(), label: label, before: before, extra: extra || null, when: new Date() }); renderLog(); }
  function letterToCol(s) { s = String(s || '').trim().toUpperCase(); if (!/^[A-Z]{1,3}$/.test(s)) return null; var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function overview() {
    var o = { loaded: !!st.grid, file: st.srcName || null, step: st.step };
    if (!st.src) { o.hint = 'Nothing is loaded yet. The person can choose a file or enter rows. If they give you numbers, type them in with enter_rows.'; return o; }
    o.sheets = st.src.worksheets.map(function (w) { return w.name; }); o.sheet = st.sheetName; o.headerRowOverride = st.headerRow || null;
    o.sheetSize = st.grid ? { rows: st.grid.maxR, cols: E.colLetter(st.grid.maxC) } : null;
    if (st.parsed) o.tables = st.parsed.blocks.map(function (b) {
      var cols = {}; Object.keys(b.col).forEach(function (k) { if (b.col[k] != null) cols[E.ROLE_LABEL[k]] = E.colLetter(b.col[k]) + ' (' + String(st.grid.cells[b.headerRow + ',' + b.col[k]] || '').slice(0, 50) + ')'; });
      return { title: b.title, headerRow: b.headerRow, lastRow: b.endRow, columns: cols, gradeRows: b.rows.length, totalRows: b.totals.length };
    });
    var c = st.cfg;
    o.rules = { acceptableFromPct: c.acc * 100, goodFromPct: c.good * 100, veryGoodFromPct: c.vg * 100, outstandingFromPct: c.out * 100, veryWeakBelowPct: c.vw == null ? null : c.vw * 100,
      weakLimitPct: c.capWeak * 100, outstandingLimitPct: c.capOut * 100, phaseTotalMethod: c.totalMethod, rounding: c.rounding, counts: c.basis, useAssessed: c.useAssessed, veryWeakCountsAsWeak: c.vwCountsAsWeak, substantialPoints: c.substantialPts, points: c.points, simulatedSheetTitle: c.simTitle,
      extraRules: (c.extraRules || []).map(function (r, i) { return { index: i + 1, rule: r.label || E.ruleLabel(r) }; }) };
    o.simulationLimits = { maxStudentsPerGrade: c.simMaxStudents, maxPointsPerGrade: c.simMaxPts, gradesLeftAlone: [], ownChanges: [] };
    if (st.base) {
      var gById = {};
      o.units = st.base.units.map(function (u) { return { table: u.block, group: u.group, phase: u.phase, grades: u.grades.map(function (g) { gById[g.id] = g; return { grade: g.grade, row: g.r, students: g.E, assessed: g.A, expectedOrBetter: g.G, betterThanExpected: g.H, workedOutFromPercentages: !!g.derived || undefined, judgement: g.judg }; }), total: u.total ? { judgement: u.total.judg, average: u.total.avg, roundingToConfirm: u.total.roundingFlag } : null }; });
      (c.simLocked || []).forEach(function (id) { if (gById[id]) o.simulationLimits.gradesLeftAlone.push({ grade: gById[id].grade, row: gById[id].r }); });
      Object.keys(c.simPinned || {}).forEach(function (id) { var g = gById[id]; if (g) o.simulationLimits.ownChanges.push({ grade: g.grade, row: g.r, expectedOrBetter: c.simPinned[id].G, betterThanExpected: c.simPinned[id].H }); });
      var k = E.check(E.counted(st.base.items, c.basis), c);
      o.actual = { counted: k.n, weak: k.weak, weakAllowed: k.maxW, outstanding: k.out, outstandingAllowed: k.maxO, extraRules: k.extra.map(function (x) { return x.label + ': ' + x.count + (x.ok ? ' (met)' : ' (not met)'); }) };
      o.fixes = E.findFixes(st.parsed, c).fixes.map(function (f) { return f.what + ': ' + f.detail; });
      if (st.sim) o.simulation = !st.sim.needed ? 'not needed' : st.sim.failed ? (st.sim.limited ? 'not possible inside the simulation limits' : 'not possible with these rules') : st.sim.options.map(function (op, i) { return { option: String.fromCharCode(65 + i), selected: i === st.sel, students: op.students, judgementChanges: op.changed.length, largestChangePoints: +op.maxPts.toFixed(1), grades: op.cells.map(function (x) { return x.to.grade + ' (row ' + x.to.r + ') ' + (x.dG ? (x.dG > 0 ? '+' : '') + x.dG + ' expected' : '') + (x.dH ? ' ' + (x.dH > 0 ? '+' : '') + x.dH + ' better' : ''); }), extraRulesNotMet: op.check.extra.filter(function (x) { return !x.ok; }).map(function (x) { return x.label; }) }; });
    } else if (st.grid) o.problem = 'No progress table was recognised with the current settings.';
    return o;
  }
  function gradeAt(row) { var it = st.base && st.base.items.filter(function (x) { return x.kind === 'grade' && x.r === +row; })[0]; if (!it) throw new Error('No grade on row ' + row + '. Use the row numbers from get_overview.'); return it; }
  var CAT_SETS = { 'Outstanding': ['Outstanding'], 'Very Good or better': ['Outstanding', 'Very Good'], 'Good or better': ['Outstanding', 'Very Good', 'Good'], 'Acceptable or better': ['Outstanding', 'Very Good', 'Good', 'Acceptable'], 'Very Good': ['Very Good'], 'Good': ['Good'], 'Acceptable': ['Acceptable'], 'Weak or worse': ['Weak', 'Very Weak'], 'Weak': ['Weak'], 'Very Weak': ['Very Weak'] };
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
    { name: 'set_cells', description: 'Changes several cells at once in the working copy of the uploaded sheet (same rules as set_cell). Use it to rearrange a layout in one go. At most 200 cells.',
      inputSchema: { type: 'object', properties: { cells: { type: 'array', items: { type: 'object', properties: { cell: { type: 'string' }, value: { type: ['string', 'number', 'null'] } }, required: ['cell', 'value'] } } }, required: ['cells'] },
      execute: function (i) {
        var list = (i.cells || []).slice(0, 200); if (!list.length) throw new Error('No cells given');
        if (!st.src) throw new Error('Nothing is loaded. Use enter_rows instead.');
        var ws = st.src.getWorksheet(st.sheetName), b = snapshot(), prev = [];
        list.forEach(function (x) { var m = /^([A-Z]+)(\d+)$/i.exec(String(x.cell).trim()); if (!m) throw new Error('Bad cell ' + x.cell); });
        list.forEach(function (x) { var ref = String(x.cell).trim().toUpperCase(), cell = ws.getCell(ref), v = x.value; if (typeof v === 'string' && v.trim() !== '' && !isNaN(+v)) v = +v; prev.push({ cell: ref, prev: cell.value }); cell.value = v; });
        st.grid = gridFromWs(ws); recompute();
        record(list.length + ' cells changed in the sheet', b, { cells: prev, sheet: st.sheetName }); return overview().tables || [];
      } },
    { name: 'enter_rows', description: 'Types rows of numbers into the page as a new table, replacing what is loaded (undo brings it back). Use it when the person gives you numbers in the chat, or when a sheet cannot be mapped any other way and you have read the numbers with read_cells. kind "counts": expected and better are numbers of students. kind "pct": they are percentages (0 to 100). Never make up numbers.',
      inputSchema: { type: 'object', properties: { title: { type: 'string' }, kind: { type: 'string', enum: ['counts', 'pct'] }, rows: { type: 'array', items: { type: 'object', properties: { phase: { type: ['string', 'number'] }, grade: { type: 'string' }, students: { type: 'number' }, expected: { type: 'number' }, better: { type: 'number' } }, required: ['grade', 'students', 'expected', 'better'] } } }, required: ['kind', 'rows'] },
      execute: function (i) {
        var rows = (i.rows || []).slice(0, 200); if (!rows.length) throw new Error('No rows given');
        var before = snapshot(), prevSrc = { src: st.src, origWb: st.origWb, srcName: st.srcName, sheetName: st.sheetName, grid: st.grid, rows: JSON.parse(JSON.stringify(ROWS)) };
        ROWS.kind = i.kind === 'pct' ? 'pct' : 'counts'; ROWS.title = String(i.title || '').slice(0, 80);
        ROWS.list = rows.map(function (r) { return { phase: r.phase == null ? '' : String(r.phase), grade: String(r.grade), E: String(r.students), G: String(r.expected), H: String(r.better) }; });
        var res = useRows('rows typed in by the assistant');
        if (!res.ok) { ROWS.kind = prevSrc.rows.kind; ROWS.title = prevSrc.rows.title; ROWS.list = prevSrc.rows.list; throw new Error('Rows not accepted: ' + res.problems.join('; ')); }
        if (!$('rows-card').hidden) renderRows();
        record('Typed in ' + res.rows + ' row' + (res.rows === 1 ? '' : 's') + (ROWS.title ? ' for ' + ROWS.title : ''), before, { source: prevSrc });
        return overview();
      } },
    { name: 'set_phase', description: 'Sets the phase of one or more grades (for example when the sheet has no phase column). Identify grades by row number from get_overview.',
      inputSchema: { type: 'object', properties: { grades: { type: 'array', items: { type: 'object', properties: { row: { type: 'integer' }, phase: { type: ['string', 'number'] } }, required: ['row', 'phase'] } } }, required: ['grades'] },
      execute: function (i) {
        var b = snapshot(), done = [];
        (i.grades || []).forEach(function (x) { var g = gradeAt(x.row); (st.edits[g.id] = st.edits[g.id] || {}).phase = String(x.phase); done.push(g.grade); });
        if (!done.length) throw new Error('No grades given');
        recompute(); record('Phase set for ' + done.join(', '), b); return overview().units;
      } },
    { name: 'add_rule', description: 'Adds an extra rule the results and simulation are checked against, for example at least 50% Good or better. op "min" (at least) or "max" (no more than). judgements: one of ' + Object.keys(CAT_SETS).join(', ') + '. Only add rules the person asks for.',
      inputSchema: { type: 'object', properties: { op: { type: 'string', enum: ['min', 'max'] }, percent: { type: 'number' }, judgements: { type: 'string', enum: Object.keys(CAT_SETS) } }, required: ['op', 'percent', 'judgements'] },
      execute: function (i) {
        var cats = CAT_SETS[i.judgements]; if (!cats) throw new Error('Unknown judgements');
        var pct = Math.max(0, Math.min(100, +i.percent)), b = snapshot();
        var r = { op: i.op === 'max' ? 'max' : 'min', share: pct / 100, cats: cats.slice() };
        st.cfg.extraRules = (st.cfg.extraRules || []).concat([r]); st.sel = 0; recompute();
        record('Rule added: ' + E.ruleLabel(r), b); return overview().actual || 'Added';
      } },
    { name: 'remove_rule', description: 'Removes an extra rule by its index from get_overview (rules.extraRules).',
      inputSchema: { type: 'object', properties: { index: { type: 'integer' } }, required: ['index'] },
      execute: function (i) { var n = +i.index - 1, L = st.cfg.extraRules || []; if (!L[n]) throw new Error('No rule ' + i.index); var b = snapshot(), lab = L[n].label || E.ruleLabel(L[n]); L.splice(n, 1); st.sel = 0; recompute(); record('Rule removed: ' + lab, b); return 'Done'; } },
    { name: 'set_simulation_limits', description: 'Limits what the simulation may change. maxStudentsPerGrade and maxPointsPerGrade cap the change to any one grade (null removes the cap). leaveAlone and allowAgain take grade row numbers. reset: true clears every limit and every own change.',
      inputSchema: { type: 'object', properties: { maxStudentsPerGrade: { type: ['number', 'null'] }, maxPointsPerGrade: { type: ['number', 'null'] }, leaveAlone: { type: 'array', items: { type: 'integer' } }, allowAgain: { type: 'array', items: { type: 'integer' } }, reset: { type: 'boolean' } } },
      execute: function (i) {
        var c = st.cfg, b = snapshot(), done = [];
        if (i.reset) { c.simMaxStudents = null; c.simMaxPts = null; c.simLocked = []; c.simPinned = {}; done.push('Simulation limits cleared'); }
        if (i.maxStudentsPerGrade !== undefined) { c.simMaxStudents = i.maxStudentsPerGrade == null ? null : Math.max(0, Math.round(+i.maxStudentsPerGrade)); done.push(c.simMaxStudents == null ? 'No cap on students per grade' : 'At most ' + c.simMaxStudents + ' students per grade'); }
        if (i.maxPointsPerGrade !== undefined) { c.simMaxPts = i.maxPointsPerGrade == null ? null : Math.max(0, +i.maxPointsPerGrade); done.push(c.simMaxPts == null ? 'No cap on points per grade' : 'At most ' + c.simMaxPts + ' points per grade'); }
        (i.leaveAlone || []).forEach(function (r) { var g = gradeAt(r); if ((c.simLocked || []).indexOf(g.id) < 0) c.simLocked = (c.simLocked || []).concat([g.id]); done.push(g.grade + ' left alone'); });
        (i.allowAgain || []).forEach(function (r) { var g = gradeAt(r); c.simLocked = (c.simLocked || []).filter(function (x) { return x !== g.id; }); done.push(g.grade + ' may change again'); });
        if (!done.length) throw new Error('Nothing to change');
        st.sel = 0; recompute(); record(done.join('. '), b); return overview().simulation;
      } },
    { name: 'set_simulated_numbers', description: 'Makes a change of the person\'s own choosing in the simulated scenario only (actual results stay as they are). The search then adds the fewest further changes needed. Give expected and/or better as student counts; null removes your change for that grade. Counts must stay whole, within assessed students, and better cannot exceed expected.',
      inputSchema: { type: 'object', properties: { row: { type: 'integer' }, expected: { type: ['integer', 'null'] }, better: { type: ['integer', 'null'] } }, required: ['row'] },
      execute: function (i) {
        var g = gradeAt(i.row), c = st.cfg, b = snapshot(); c.simPinned = c.simPinned || {};
        if (i.expected === null && i.better === null) { delete c.simPinned[g.id]; st.sel = 0; recompute(); record('Own simulated change removed for ' + g.grade, b); return overview().simulation; }
        var p = Object.assign({}, c.simPinned[g.id] || {});
        if (i.expected != null) p.G = Math.round(+i.expected); if (i.better != null) p.H = Math.round(+i.better);
        var G = p.G != null ? p.G : g.G, H = p.H != null ? p.H : g.H;
        if (G < 0 || H < 0 || G > g.A || H > g.A) throw new Error('Counts must be between 0 and ' + g.A + ' assessed students');
        if (H > G) throw new Error('Better than expected cannot be more than expected or better');
        c.simPinned[g.id] = p; st.sel = 0; recompute();
        record('Simulated ' + g.grade + ': ' + (p.G != null ? 'expected ' + g.G + ' → ' + p.G : '') + (p.H != null ? (p.G != null ? ', ' : '') + 'better ' + g.H + ' → ' + p.H : ''), b);
        return overview().simulation;
      } },
    { name: 'report_gap', description: 'Tells the site owner about something you could not do: a change the tools cannot make, a missing feature, or a sheet layout you could not read. Call it once per gap, after telling the person you cannot do it. Write a short general summary. Never include student numbers, names, cell values or anything copied from the sheet.',
      inputSchema: { type: 'object', properties: { asked: { type: 'string', description: 'What the person wanted, in one or two general sentences.' }, why: { type: 'string', description: 'Why the tools could not do it, and what would be needed.' }, kind: { type: 'string', enum: ['missing feature', 'layout not readable', 'bug', 'other'] } }, required: ['asked', 'why'] },
      execute: function (i) {
        if (AS.reported >= 3) return 'Already reported several things in this visit; not sending more.';
        return fetch('/api/report', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ asked: String(i.asked || '').slice(0, 400), why: String(i.why || '').slice(0, 600), kind: i.kind || 'missing feature', step: 'step ' + st.step }) })
          .then(function (r) { return r.json().catch(function () { return {}; }).then(function (j) { if (!r.ok) throw new Error(j.error || 'Could not send the report'); AS.reported = (AS.reported || 0) + 1; bubble('note', 'Passed on to the site owner.'); return 'Reported to the site owner' + (j.number ? ' (#' + j.number + ')' : '') + '.'; }); });
      } },
    { name: 'go_to_step', description: 'Shows a step of the page so the person can see the result: 1 Load, 2 Check, 3 Rules (including extra rules), 4 Results, 5 Simulation (including its limits), 6 Download.', inputSchema: { type: 'object', properties: { step: { type: 'integer' } }, required: ['step'] },
      execute: function (i) { go(+i.step); if (+i.step === 5) { var d = $('simlim'); if (d && (st.cfg.simMaxStudents != null || st.cfg.simMaxPts != null || (st.cfg.simLocked || []).length || Object.keys(st.cfg.simPinned || {}).length)) d.open = true; } return 'Showing step ' + st.step; } }
  ];

  /* ---------- transports ---------- */
  var PRIORITY = ['get_overview', 'read_cells', 'set_column', 'set_header_row', 'set_cells', 'enter_rows', 'set_rules', 'add_rule', 'set_simulation_limits', 'set_simulated_numbers', 'choose_option', 'go_to_step', 'edit_number', 'set_phase', 'remove_rule', 'choose_sheet', 'set_cell'];
  function detect() {
    if (window.claude && typeof window.claude.use === 'function') {
      return window.claude.use('sample').then(function (s) {
        if (!s) return null;
        AS.sample = s;
        return (s.limits ? s.limits().catch(function () { return null; }) : Promise.resolve(null)).then(function (l) {
          if (l && l.tools) { AS.maxTools = l.tools.maxCount || TOOLS.length; return 'claude'; }
          return 'claude-text';
        });
      }).catch(function () { return null; });
    }
    return fetch('/api/health', { method: 'GET' }).then(function (r) { return r.ok ? r.json() : null; }).then(function (j) { AS.canReport = !!(j && j.reports); return j && j.ok ? 'api' : null; }).catch(function () { return null; });
  }
  function pageTools() {
    var max = Math.max(1, AS.maxTools || TOOLS.length);
    var list = TOOLS.filter(function (t) { return t.name !== 'report_gap' || AS.canReport; }).sort(function (a, b) { var x = PRIORITY.indexOf(a.name), y = PRIORITY.indexOf(b.name); return (x < 0 ? 99 : x) - (y < 0 ? 99 : y); }).slice(0, max);
    return list.map(function (t) { return { name: t.name, description: t.description, inputSchema: t.inputSchema, execute: function (inp) { activity(t.name, inp); return t.execute(inp || {}); } }; });
  }
  function runClaude(userText, onText) {
    var turns = [{ role: 'user', content: INSTRUCTIONS + '\n\nCURRENT STATE (JSON):\n' + JSON.stringify(overview()).slice(0, 60000) }].concat(AS.turns.slice(-12)).concat([{ role: 'user', content: userText }]);
    var noTools = AS.mode === 'claude-text';
    if (noTools) turns[0].content += '\n\nIN THIS VIEW YOU HAVE NO TOOLS: you can explain and advise, but you cannot change anything on the page. If they ask for a change, tell them exactly where on the page to make it (step and field).';
    return AS.sample(turns, noTools ? { cache: false, signal: AS.ctl.signal, onText: function (u) { onText(u.text); } } : { tools: pageTools(), cache: false, signal: AS.ctl.signal, onText: function (u) { onText(u.text); } })
      .then(function (r) { return r.text; });
  }
  function runApi(userText, onText) {
    var msgs = AS.turns.slice(-12).map(function (t) { return { role: t.role, content: t.content }; });
    msgs.push({ role: 'user', content: 'CURRENT STATE (JSON):\n' + JSON.stringify(overview()).slice(0, 60000) + '\n\nREQUEST:\n' + userText });
    var tools = TOOLS.filter(function (t) { return t.name !== 'report_gap' || AS.canReport; }).map(function (t) { return { name: t.name, description: t.description, input_schema: t.inputSchema }; });
    var said = '';
    function round(n) {
      if (n > 12) return Promise.resolve(said || 'I stopped after several steps. Check the changes listed below.');
      return fetch('/api/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ messages: msgs, tools: tools }), signal: AS.ctl.signal })
        .then(function (r) { if (!r.ok) return r.json().catch(function () { return {}; }).then(function (j) { throw { code: 'api', message: j.error || ('The assistant is unavailable (' + r.status + ').') }; }); return r.json(); })
        .then(function (res) {
          var blocks = res.content || [], text = blocks.filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('\n');
          if (text) { said += (said ? '\n\n' : '') + text; onText(said); }
          msgs.push({ role: 'assistant', content: blocks });
          var uses = blocks.filter(function (b) { return b.type === 'tool_use'; });
          if (res.stop_reason !== 'tool_use' || !uses.length) return said;
          return Promise.all(uses.map(function (u) {
            var t = TOOLS.filter(function (x) { return x.name === u.name; })[0];
            function wrap(out, err) { return { type: 'tool_result', tool_use_id: u.id, content: typeof out === 'string' ? out : JSON.stringify(out === undefined ? 'Done' : out).slice(0, 30000), is_error: err }; }
            return Promise.resolve().then(function () { if (!t) throw new Error('Unknown tool'); activity(u.name, u.input); return t.execute(u.input || {}); })
              .then(function (out) { return wrap(out, false); }, function (e) { return wrap('Error: ' + (e && e.message || e), true); });
          })).then(function (results) {
            msgs.push({ role: 'user', content: results });
            return round(n + 1);
          });
        });
    }
    return round(1);
  }

  /* ---------- pop-up UI ---------- */
  var panel = $('asst'), log = $('asst-log'), input = $('asst-in'), fab = $('asst-open'), unseen = 0;
  function isOpen() { return panel.dataset.open === 'true'; }
  function openPanel(open) {
    panel.dataset.open = open ? 'true' : 'false'; panel.setAttribute('aria-hidden', open ? 'false' : 'true'); fab.setAttribute('aria-expanded', open ? 'true' : 'false');
    document.body.classList.toggle('asst-open', open);
    if (open) { unseen = 0; badge(); renderSugg(); setTimeout(function () { input.focus(); }, 30); var bd = $('asst-body'); bd.scrollTop = bd.scrollHeight; }
    else if (panel.contains(document.activeElement)) fab.focus();
  }
  function badge() { var b = $('asst-badge'); b.hidden = !unseen; b.textContent = unseen; }
  fab.addEventListener('click', function () { openPanel(!isOpen()); });
  $('asst-close').addEventListener('click', function () { openPanel(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && isOpen()) openPanel(false); });
  function bubble(role, text, cls) { var d = document.createElement('div'); d.className = 'bub ' + role + (cls ? ' ' + cls : ''); d.textContent = text; log.appendChild(d); $('asst-hello').hidden = true; var bd = $('asst-body'); bd.scrollTop = bd.scrollHeight; return d; }
  var NICE = { get_overview: 'Looking at the page', read_cells: 'Reading cells', choose_sheet: 'Switching sheet', set_header_row: 'Setting the heading row', set_column: 'Matching a column', set_rules: 'Changing rules', edit_number: 'Correcting a number', set_cell: 'Fixing a cell', set_cells: 'Fixing cells', enter_rows: 'Typing the rows in', set_phase: 'Setting phases', add_rule: 'Adding a rule', remove_rule: 'Removing a rule', set_simulation_limits: 'Setting simulation limits', set_simulated_numbers: 'Changing the simulation', report_gap: 'Passing it on', choose_option: 'Choosing an option', go_to_step: 'Opening a step' };
  var actEl = null;
  function activity(name) { if (actEl) actEl.textContent = (NICE[name] || name) + '…'; }
  function renderSugg() {
    var q;
    if (!st.grid) q = [['f', 'Here are my numbers: '], ['s', 'What does this page do?'], ['s', 'What should my sheet look like?']];
    else if (!st.base) q = [['s', "My sheet wasn't read properly. Please fix it."], ['f', 'The headings are in row ']];
    else if (st.step === 5) q = [['s', 'Explain this simulation simply.'], ['f', 'Leave this grade alone in the simulation: '], ['f', 'Add at most this many students to any grade: '], ['s', 'Show me the other options.']];
    else if (st.step === 3) q = [['f', 'Add a rule: at least 50% Good or better'], ['s', 'Use round down for phase averages.'], ['s', 'Only count grade judgements toward the limits.']];
    else q = [['s', 'Explain the results simply.'], ['f', 'Add a rule: '], ['f', 'Change the simulation so '], ['s', 'Check the sheet was read correctly.']];
    $('asst-sugg').innerHTML = q.map(function (x) { return '<button type="button" class="chip-btn" data-' + (x[0] === 's' ? 'q' : 'fill') + '="' + esc(x[1]) + '">' + esc(x[1].replace(/:\s*$|\s+$/, '')) + '</button>'; }).join('');
  }
  function renderLog() {
    var box = $('asst-changes');
    box.hidden = !AS.log.length;
    $('asst-changes-t').textContent = 'What it changed (' + AS.log.length + ')';
    $('asst-changes-list').innerHTML = AS.log.map(function (x, i) { return '<li><span class="dot sim" aria-hidden="true"></span><span>' + esc(x.label) + '</span>' + (i === 0 ? '<button type="button" class="btn quiet" data-undo="' + x.id + '">Undo</button>' : '<span></span>') + '</li>'; }).join('');
    var banner = $('asst-banner');
    if (banner) { banner.hidden = !AS.log.length; $('asst-banner-t').textContent = AS.log.length ? 'The assistant made ' + AS.log.length + ' change' + (AS.log.length === 1 ? '' : 's') + '. Latest: ' + AS.log[0].label + '.' : ''; }
    if (!isOpen() && AS.log.length) { unseen++; badge(); }
  }
  document.addEventListener('click', function (e) {
    var u = e.target.closest && e.target.closest('[data-undo]'); if (!u) return;
    var x = AS.log[0]; if (!x) return;
    if (x.extra && x.extra.cell) { var ws = st.src.getWorksheet(x.extra.sheet); if (ws) ws.getCell(x.extra.cell).value = x.extra.prev; }
    if (x.extra && x.extra.cells) { var ws2 = st.src.getWorksheet(x.extra.sheet); if (ws2) x.extra.cells.slice().reverse().forEach(function (p) { ws2.getCell(p.cell).value = p.prev; }); }
    AS.log.shift();
    if (x.extra && x.extra.source) {
      var ps = x.extra.source; ROWS.kind = ps.rows.kind; ROWS.title = ps.rows.title; ROWS.list = ps.rows.list; if (!$('rows-card').hidden) renderRows();
      st.src = ps.src; st.origWb = ps.origWb; st.srcName = ps.srcName; st.sheetName = ps.sheetName; st.grid = ps.grid;
      if (!st.src) { st.parsed = null; st.base = null; st.sim = null; st.cfg = x.before.cfg; st.edits = {}; renderLog(); go(1); bubble('note', 'Undone: ' + x.label + '.'); return; }
    }
    restore(x.before); renderLog(); bubble('note', 'Undone: ' + x.label + '.');
  });
  $('asst-banner-open') && $('asst-banner-open').addEventListener('click', function () { openPanel(true); $('asst-changes').open = true; });
  function grow() { input.style.height = 'auto'; input.style.height = Math.min(140, input.scrollHeight + 2) + 'px'; }
  input.addEventListener('input', grow);
  function send(text) {
    text = String(text || '').trim(); if (!text || AS.busy) return;
    if (AS.mode === undefined) { bubble('note', 'Still connecting to the assistant. Try again in a moment.'); return; }
    if (!AS.mode) { bubble('note', UNAVAILABLE); return; }
    AS.busy = true; AS.ctl = new AbortController(); $('asst-send').hidden = true; $('asst-stop').hidden = false;
    bubble('user', text); input.value = ''; grow();
    var out = bubble('bot', ''); actEl = document.createElement('div'); actEl.className = 'act'; actEl.textContent = 'Thinking…'; out.appendChild(actEl);
    var p = AS.mode === 'claude' || AS.mode === 'claude-text' ? runClaude(text, function (t) { out.textContent = t; out.appendChild(actEl); }) : runApi(text, function (t) { out.textContent = t; out.appendChild(actEl); });
    p.then(function (final) {
      out.textContent = final || 'Done.'; AS.turns.push({ role: 'user', content: text }, { role: 'assistant', content: final || 'Done.' });
    }, function (e) {
      var c = e && e.code;
      if (c === 'cancelled' || (e && e.name === 'AbortError')) out.textContent = (e.text || out.textContent.replace(/\S*…$/, '')) + ' (stopped)';
      else if (c === 'not_granted' || c === 'sampling_disabled') { out.textContent = 'The assistant needs permission to use Claude, and it was not given.'; }
      else if (c === 'rate_limited') out.textContent = 'Too many requests right now. Try again in a little while.';
      else out.textContent = (e && e.message) || 'Something went wrong. Try again.';
      out.classList.add('err');
    }).then(function () { AS.busy = false; $('asst-send').hidden = false; $('asst-stop').hidden = true; actEl = null; renderSugg(); var bd = $('asst-body'); bd.scrollTop = bd.scrollHeight; if (isOpen()) input.focus(); });
  }
  openAssistantWith = function (q) { openPanel(true); send(q); };
  $('asst-form').addEventListener('submit', function (e) { e.preventDefault(); send(input.value); });
  input.addEventListener('keydown', function (e) { if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); send(input.value); } });
  $('asst-stop').addEventListener('click', function () { if (AS.ctl) AS.ctl.abort(); });
  $('asst-sugg').addEventListener('click', function (e) {
    var b = e.target.closest('[data-q],[data-fill]'); if (!b) return;
    if (b.dataset.q) send(b.dataset.q); else { input.value = b.dataset.fill; grow(); input.focus(); input.setSelectionRange(input.value.length, input.value.length); }
  });
  var UNAVAILABLE = window.claude ? 'The assistant cannot reach Claude in this view. Open this page in Claude on a computer (web or desktop app) and allow it to use Claude when asked. Everything else on the page works here.' : 'The assistant is not switched on for this site yet. Everything else on the page works.';
  AS.mode = undefined;
  fab.hidden = false;
  $('asst-where').textContent = 'Connecting…';
  detect().then(function (m) {
    AS.mode = m;
    $('asst-where').textContent = m === 'claude' ? 'Uses Claude on your own account. What it needs from your sheet is sent only when you ask.'
      : m === 'claude-text' ? 'In this view the assistant can explain and advise, but cannot change the page. Open it in Claude on a computer to let it make changes.'
      : m === 'api' ? 'Uses Claude through this site. What it needs from your sheet is sent only when you ask, and is not stored.' : '';
    if (!m) { $('asst-hello').innerHTML = '<b>Assistant unavailable here</b>' + esc(UNAVAILABLE); $('asst-sugg').hidden = true; $('asst-in').disabled = true; $('asst-send').disabled = true; }
    renderSugg();
  });
