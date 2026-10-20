var View = (function (C) {
  function validate(rows) {
    return rows.map(function (r) {
      var e = [];
      if (r.E == null && r.G == null && r.H == null) return ['empty'];
      if (!(r.E > 0)) e.push('Students must be more than 0.');
      if (r.G == null || r.G < 0 || r.G > r.E) e.push('Expected or better must be between 0 and the number of students.');
      if (r.H == null || r.H < 0 || r.H > r.G) e.push('Better than expected can’t be more than expected or better.');
      if (r.phase == null || r.phase === '') e.push('Add a phase number.');
      if (!r.grade) e.push('Add a grade name.');
      return e;
    });
  }
  function editorRows(rows, errs) {
    return rows.map(function (r, i) {
      var g = C.esc(r.grade || ''), bad = errs && errs[i] && errs[i].length && errs[i][0] !== 'empty';
      function inp(field, cls, type) {
        var v = r[field] == null ? '' : r[field];
        return '<td class="' + cls + '"><input type="' + type + '" id="e-' + field + '-' + i + '" name="e-' + field + '-' + i + '" data-i="' + i + '" data-f="' + field + '" value="' + C.esc(v) + '"' +
          (type === 'number' ? ' min="0" step="1" inputmode="numeric"' : ' spellcheck="false"') + ' autocomplete="off" aria-label="' + C.esc(({ phase: 'Phase', grade: 'Grade', E: 'Students', G: 'Made expected or better', H: 'Did better than expected' })[field] + ', row ' + (i + 1)) + '"' +
          (bad && field !== 'grade' && field !== 'phase' ? ' aria-invalid="true"' : '') + '></td>';
      }
      return '<tr>' + inp('phase', 'ph', 'number') + inp('grade', 'gr', 'text') + inp('E', '', 'number') + inp('G', '', 'number') + inp('H', '', 'number') +
        '<td class="rm"><button type="button" class="btn btn-quiet" data-rm="' + i + '" aria-label="Remove ' + C.esc(C.gradeName(r.grade || 'row ' + (i + 1))) + '">Remove</button></td></tr>' +
        (bad ? '<tr class="err"><td></td><td colspan="5" class="err-msg">' + C.esc(errs[i].join(' ')) + ' This row is left out until fixed.</td></tr>' : '');
    }).join('');
  }
  function stats(model, sim, r) {
    var list = C.counted(model.items, r.mode), d = C.dist(list);
    var w = d.counts['Weak'], maxW = Math.floor(r.capWeak * d.n + 1e-9);
    var over = Math.max(0, w - maxW);
    var fix;
    if (!sim.needed) fix = '<dd class="ok">Nothing<small>it already fits</small></dd>';
    else if (sim.failed) fix = '<dd class="bad">Not possible<small>by adding students alone</small></dd>';
    else {
      var o = sim.options[0];
      var g = o.cellChanges.filter(function (c) { return c.from.kind === 'grade'; }).length;
      fix = '<dd class="sim">' + o.students + ' students<small>more making progress, in ' + g + (g === 1 ? ' grade' : ' grades') + '</small></dd>';
    }
    return '<div><dt>Weak right now</dt><dd class="' + (over ? 'bad' : 'ok') + '">' + w + ' of ' + d.n + '<small>' + ({ all: 'grades and phase totals', grades: 'grades', totals: 'phase totals' })[r.mode] + '</small></dd></div>' +
      '<div><dt>Weak allowed</dt><dd>' + maxW + '<small>' + C.pct(r.capWeak, 0) + ' of ' + d.n + '</small></dd></div>' +
      '<div><dt>Too many</dt><dd class="' + (over ? 'bad' : 'ok') + '">' + over + '<small>' + (over ? 'need to stop being Weak' : 'within the limit') + '</small></dd></div>' +
      '<div><dt>Smallest pretend fix</dt>' + fix + '</div>';
  }
  function all(parsed, rows, r, sel, wb1) {
    wb1 = wb1 || 0;
    var subj = (parsed && parsed.subject) || 'Your sheet';
    var errs = validate(rows);
    var good = rows.filter(function (x, i) { return !errs[i].length; });
    if (!good.length) return { empty: true, sel: 0, EDITOR: editorRows(rows, errs), N_ALL: '0', N_GRADES: '0', N_TOTALS: '0' };
    var model = C.build(good, r);
    var sim = C.simulate(good, r);
    if (sim.options && sel >= Math.min(sim.options.length, 6)) sel = 0;
    var opt = sim.options && sim.options.length ? sim.options[sel] : null;
    var nG = model.items.filter(function (x) { return x.kind === 'grade'; }).length, nT = model.items.length - nG;
    return {
      sel: sel, sim: sim, model: model,
      STATS: stats(model, sim, r),
      EDITOR: editorRows(rows, errs),
      N_ALL: String(model.items.length), N_GRADES: String(nG), N_TOTALS: String(nT),
      DIST_NOW: C.renderDist(C.counted(model.items, r.mode), r),
      TABLE_NOW: C.renderTable(model, r),
      FIXES: C.renderFixes(parsed, r),
      COMPARE: C.renderCompare(parsed, model, opt, r),
      SIM_HEAD: C.renderSimHeader(sim, r),
      SIM_TABS: sim.options ? C.renderTabs(sim, sel) : '',
      WB1_TABS: C.renderSheetTabs(['Original (as received)', 'Fixed (real numbers)'], wb1, 'wtab', 'wb1-grid'),
      WB1_GRID: !(parsed && parsed.raw && parsed.raw.length) ? '' : wb1 === 0 ? C.renderGrid(C.sheetSpecRaw(parsed, subj), { label: 'Original sheet' }) : C.renderGrid(C.sheetSpecModel(model, r, null, subj + ' (fixed)'), { label: 'Fixed sheet' }),
      WB1_NOTES: !(parsed && parsed.raw && parsed.raw.length) ? '' : wb1 === 0 ? (function () { var n = C.sheetSpecRaw(parsed).notes; return n.length ? '<ul class="xl-notes">' + n.map(function (x) { return '<li>' + C.esc(x) + '</li>'; }).join('') + '</ul>' : ''; })() : '<p class="xl-notes-p">Same students, with the totals added up properly and the phase labels in the right place. Point at a cell in the Original tab to see what was wrong.</p>',
      WB2_TABS: sim.options && sim.options.length ? C.renderSheetTabs(sim.options.slice(0, 6).map(function (o, i) { return 'Option ' + C.letter(i) + (i === 0 ? ' (fewest)' : ''); }), sel, 'otab', 'wb2-grid') : '',
      WB2_GRID: opt ? C.renderGrid(C.sheetSpecModel(opt.model, r, model, subj + ' (SIMULATED, Option ' + C.letter(sel) + ')'), { label: 'Simulated sheet, option ' + C.letter(sel) }) : '',
      WB2_CHANGES: opt ? '<p class="small"><strong>' + opt.students + ' more students</strong> would need to make expected progress. ' + opt.changed.length + ' results change:</p>' + C.renderChangeList(sim, sel) : '',
      SHEET_RAW: '',
      SHEET_TABS: sim.options ? C.renderTabs(sim, sel, 'stab') : '',
      SHEET_TITLE: opt ? 'Option ' + C.letter(sel) + ': ' + opt.students + ' more students, ' + opt.changed.length + ' results change' : '',
      SHEET_BODY: opt ? C.renderSimSheet(sim, sel, r) : '',
      SHEET_NONE: sim.needed ? (sim.failed ? '<p class="lead-bad"><span class="dot bad" aria-hidden="true"></span>No simulation can meet the limits by adding students alone.</p>' : '') : '<p class="lead-ok"><span class="dot ok" aria-hidden="true"></span>The real results already fit, so there is nothing to simulate.</p>',
      OPT_TITLE: opt ? 'Option ' + C.letter(sel) : '',
      OPT_BODY: opt ? C.renderOption(sim, sel, r, parsed) : '',
      PANEL_HIDDEN: opt ? '' : 'hidden'
    };
  }
  return { all: all, validate: validate };
})(Core);
if (typeof module !== 'undefined') module.exports = View;
