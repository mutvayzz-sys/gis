/* Builds formatted Excel workbooks in the browser with ExcelJS. Nothing leaves the page. */
var XLWriter = (function (C) {
  function argb(h) { return 'FF' + String(h).replace('#', '').toUpperCase(); }
  var T = { style: 'thin', color: { argb: 'FF000000' } }, M = { style: 'medium', color: { argb: 'FF000000' } };
  var YEL = 'FFFF00', CREAM = 'FEF2CB', PALE = 'FEFFB3', BLUE = 'DFE3FF', INK = '0000B8';
  function fill(h) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(h) } }; }
  function font(o) { o = o || {}; var f = { name: 'Calibri', size: o.size || 10, bold: !!o.bold, italic: !!o.italic }; if (o.color) f.color = { argb: argb(o.color) }; return f; }
  function setup(ws) { ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 }; }
  function judgeFormula(r, rules) {
    return 'IF(I' + r + '<' + rules.acc + ',"Weak",IF(J' + r + '>=' + rules.out + ',"Outstanding",IF(J' + r + '>=' + rules.vg + ',"Very Good",IF(J' + r + '>=' + rules.good + ',"Good","Acceptable"))))';
  }

  // One sheet in the file's own layout. opt: {formulas, pct, banner, bannerColor, sim, rules, tab}
  function specSheet(wb, name, spec, opt) {
    opt = opt || {};
    var ws = wb.addWorksheet(name.slice(0, 31), { properties: opt.tab ? { tabColor: { argb: argb(opt.tab) } } : {} });
    setup(ws);
    var widths = [2.7, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 15, 11];
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    var lastCol = opt.sim ? 13 : 11;
    if (opt.banner) {
      ws.mergeCells(1, 2, 1, lastCol);
      var b = ws.getCell(1, 2); b.value = opt.banner; b.font = font({ size: 11, bold: true, color: 'FFFFFF' }); b.fill = fill(opt.bannerColor || 'C00000');
      b.alignment = { horizontal: 'center', vertical: 'middle' }; ws.getRow(1).height = 22;
    }
    ws.mergeCells(2, 2, 2, 11);
    var t = ws.getCell(2, 2); t.value = spec.title; t.font = font({ size: 16, bold: true }); t.fill = fill(YEL); t.alignment = { horizontal: 'center', vertical: 'middle' };
    for (var c = 2; c <= 11; c++) ws.getCell(2, c).border = { top: M, bottom: M, left: c === 2 ? M : undefined, right: c === 11 ? M : undefined };
    ws.getRow(2).height = 21;
    spec.head.forEach(function (h, i) {
      var cell = ws.getCell(3, 2 + i); cell.value = h.v; cell.font = font({ bold: true }); if (h.fill) cell.fill = fill(h.fill);
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { top: M, bottom: M, left: i === 0 ? M : T, right: i === 9 ? M : T };
    });
    ws.getRow(3).height = 78;
    var n = spec.body.length, start = 0;
    spec.body.forEach(function (row, ri) {
      var r = 4 + ri, m = (spec.meta && spec.meta[ri]) || {}, nextT = spec.meta && spec.meta[ri + 1] && spec.meta[ri + 1].isT;
      row.forEach(function (cell, ci) {
        var x = ws.getCell(r, 2 + ci);
        var bt = (ri === 0 || m.isT) ? M : T, bb = (m.isT || ri === n - 1) ? M : T;
        if (nextT && ci > 1) bb = M;
        x.border = { top: ci <= 1 ? (ri === 0 ? M : undefined) : bt, bottom: ci <= 1 ? (ri === n - 1 ? M : undefined) : bb, left: (ci === 0 || ci === 2) ? M : T, right: (ci <= 1 || ci === 9) ? M : T };
        if (cell === 'skip' || cell == null) return;
        if (cell.rs > 1) {
          ws.mergeCells(r, 2 + ci, r + cell.rs - 1, 2 + ci);
          if (ci === 0) for (var k = 1; k < cell.rs; k++) ws.getCell(r + k, 2).border = { left: M, right: M, bottom: (ri + k === n - 1 || (spec.meta[ri + k] && spec.meta[ri + k].isT && ri + k === ri + cell.rs - 1)) ? M : undefined };
        }
        var v = cell.v;
        if (cell.t === 'n' && v != null && v !== '' && !isNaN(+v)) v = +v;
        x.value = v == null ? null : v;
        var sim = cell.mark === 'sim';
        x.fill = fill(sim ? BLUE : (cell.fill || 'FFFFFF'));
        x.font = font({ bold: cell.bold || sim, color: sim ? INK : cell.color, size: cell.size && cell.size >= 16 ? 16 : 10 });
        x.alignment = { horizontal: 'center', vertical: 'middle', textRotation: cell.rot ? 90 : 0 };
        if (cell.t === 'p') x.numFmt = opt.pct || '0%';
        if (m.isT && (ci === 5 || ci === 6)) x.numFmt = '0';
        if (cell.tip) x.note = cell.tip;
      });
      if (opt.formulas) {
        var res = function (ci) { var c0 = row[ci]; return c0 && c0 !== 'skip' ? c0.v : null; };
        if (m.isT) {
          ['E', 'G', 'H'].forEach(function (L, j) { ws.getCell(L + r).value = { formula: 'SUM(' + L + (4 + start) + ':' + L + (r - 1) + ')', result: res([3, 5, 6][j]) }; });
          start = ri + 1;
        }
        ws.getCell('I' + r).value = { formula: 'IF(E' + r + '>0,G' + r + '/E' + r + ',0)', result: res(7) };
        ws.getCell('J' + r).value = { formula: 'IF(E' + r + '>0,H' + r + '/E' + r + ',0)', result: res(8) };
        ws.getCell('K' + r).value = { formula: judgeFormula(r, opt.rules), result: res(9) };
      }
      if (opt.sim) {
        var lc = ws.getCell('L' + r), mc = ws.getCell('M' + r);
        lc.value = m.baseG != null ? m.baseG : null;
        var d = m.baseG != null ? m.G - m.baseG : 0;
        mc.value = d ? d : null; mc.numFmt = '+0;-0;0';
        [lc, mc].forEach(function (x, j) {
          x.fill = fill(CREAM); x.alignment = { horizontal: 'center', vertical: 'middle' };
          x.font = font({ bold: m.isT || (j === 1 && d), color: d ? INK : (m.isT ? 'FF0000' : null) });
          x.border = { top: (ri === 0 || m.isT) ? M : T, bottom: (m.isT || ri === n - 1 || nextT) ? M : T, left: j === 0 ? M : T, right: j === 1 ? M : T };
        });
      }
    });
    if (opt.sim) {
      [['L', 'Real value (students that made expected or better progress)'], ['M', 'Pretend change']].forEach(function (p, j) {
        var x = ws.getCell(p[0] + '3'); x.value = p[1]; x.font = font({ bold: true, color: INK }); x.fill = fill(BLUE);
        x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; x.border = { top: M, bottom: M, left: j === 0 ? M : T, right: j === 1 ? M : T };
      });
    }
    ws.views = [{ state: 'frozen', ySplit: 3 }];
    return ws;
  }

  // A plain formatted table sheet
  function tableSheet(wb, name, title, sub, headers, rows, widths, o) {
    o = o || {};
    var ws = wb.addWorksheet(name, { properties: o.tab ? { tabColor: { argb: argb(o.tab) } } : {} });
    setup(ws);
    ws.getColumn(1).width = 2.7;
    widths.forEach(function (w, i) { ws.getColumn(i + 2).width = w; });
    ws.getCell('B2').value = title; ws.getCell('B2').font = font({ size: 14, bold: true });
    if (sub) { ws.getCell('B3').value = sub; ws.getCell('B3').font = font({ italic: true, color: '595959' }); }
    var r0 = 5;
    headers.forEach(function (h, i) {
      var x = ws.getCell(r0, 2 + i); x.value = h; x.font = font({ bold: true }); x.fill = fill(YEL);
      x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      x.border = { top: M, bottom: M, left: i === 0 ? M : T, right: i === headers.length - 1 ? M : T };
    });
    ws.getRow(r0).height = 30;
    rows.forEach(function (row, j) {
      var r = r0 + 1 + j, last = j === rows.length - 1, lines = 1;
      row.forEach(function (v, i) {
        var x = ws.getCell(r, 2 + i);
        x.value = v;
        var f = (o.fills && o.fills[i]) || CREAM;
        x.fill = fill(f);
        x.font = font({ bold: (o.boldLast && last) || (o.boldCol === i), color: o.colors && o.colors[i] });
        x.alignment = { horizontal: (o.left && o.left[i]) ? 'left' : 'center', vertical: 'middle', wrapText: true };
        x.border = { top: (o.boldLast && last) ? M : T, bottom: last ? M : T, left: i === 0 ? M : T, right: i === row.length - 1 ? M : T };
        if (o.fmt && o.fmt[i]) x.numFmt = o.fmt[i];
        if (typeof v === 'string') lines = Math.max(lines, Math.ceil(v.length * 1.1 / (widths[i] || 10)));
      });
      ws.getRow(r).height = Math.max(16, 13.5 * lines + 4);
    });
    return { ws: ws, firstRow: r0 + 1, lastRow: r0 + rows.length };
  }

  function criteriaSheet(wb, rules, n) {
    var p = function (x) { return Math.round(x * 1000) / 10 + '%'; };
    var maxW = Math.floor(rules.capWeak * n + 1e-9), maxO = Math.floor(rules.capOut * n + 1e-9);
    var t = tableSheet(wb, 'Criteria', 'How each result is decided', 'From "Criteria For Attainment & Progress". Lines mean "at least".',
      ['Result', 'What is checked', 'Line', 'Rule'],
      [['Weak', 'Students making expected or better progress', 'Below ' + p(rules.acc), 'Fewer than ' + p(rules.acc) + ' make expected or better progress.'],
       ['Acceptable', 'Students making expected or better progress', p(rules.acc), 'The ' + p(rules.acc) + ' line is reached, but fewer than ' + p(rules.good) + ' did better than expected.'],
       ['Good', 'Students doing better than expected', p(rules.good), 'Acceptable, and at least ' + p(rules.good) + ' did better than expected.'],
       ['Very Good', 'Students doing better than expected', p(rules.vg), 'Acceptable, and at least ' + p(rules.vg) + ' did better than expected.'],
       ['Outstanding', 'Students doing better than expected', p(rules.out), 'Acceptable, and at least ' + p(rules.out) + ' did better than expected.'],
       ['Weak limit', 'Share of all judgements counted', p(rules.capWeak), 'No more than ' + p(rules.capWeak) + ' may be Weak: at most ' + maxW + ' of ' + n + '.'],
       ['Outstanding limit', 'Share of all judgements counted', p(rules.capOut), 'No more than ' + p(rules.capOut) + ' may be Outstanding: at most ' + maxO + ' of ' + n + '.']],
      [16, 34, 11, 62], { fills: [PALE, CREAM, YEL, CREAM], left: [false, true, false, true], tab: '2E75B6' });
    return t.ws;
  }
  function correctionsSheet(wb, parsed, rules) {
    var f = C.findFixes(parsed, rules);
    var rows = f.fixes.map(function (x, i) { return [i + 1, x.what, x.detail, x.effect || 'None']; });
    if (!rows.length) rows = [[1, 'Nothing to fix', 'Every count, percentage and result in the loaded sheet already matches the rules.', 'None']];
    var t = tableSheet(wb, 'Corrections Log', 'Mistakes fixed in the sheet', 'Only calculation and layout mistakes were fixed. No student numbers were changed.',
      ['#', 'What was wrong', 'Details', 'Effect on result'], rows, [5, 30, 70, 20], { fills: [CREAM, PALE, CREAM, PALE], left: [false, true, true, false], tab: '548235' });
    if (f.fine.length) { var c = t.ws.getCell(t.lastRow + 2, 2); c.value = f.fine.join(' '); c.font = font({ italic: true, color: '375623' }); }
    return t.ws;
  }
  function changesSheet(wb, sim, idxs) {
    var rows = [];
    idxs.forEach(function (i) {
      var o = sim.options[i], L = 'Option ' + C.letter(i);
      o.cellChanges.filter(function (c) { return c.from.kind === 'grade'; }).forEach(function (c) {
        rows.push([L, c.to.grade, c.from.G, c.to.G, c.to.G - c.from.G, c.from.judg, c.to.judg]);
      });
      o.changed.filter(function (c) { return c.from.kind === 'total'; }).forEach(function (c) { rows.push([L, 'Phase ' + c.to.phase + ' total', null, null, null, c.from.judg, c.to.judg]); });
      rows.push([L + ' total', null, null, null, o.students, null, o.changed.length + ' results change']);
    });
    var t = tableSheet(wb, 'Changes', 'Every pretend change', 'SIMULATED. These are not actual results.',
      ['Option', 'Grade', 'Real value', 'Pretend value', 'Change', 'Result before', 'Result after'], rows, [16, 16, 11, 13, 10, 15, 18],
      { fills: [CREAM, CREAM, CREAM, BLUE, BLUE, PALE, PALE], fmt: [null, null, null, null, '+0;-0;0'], colors: [null, null, null, INK, INK], tab: 'C00000' });
    rows.forEach(function (r, j) { if (/total$/.test(r[0])) t.ws.getRow(t.firstRow + j).eachCell(function (c) { c.font = font({ bold: true, color: 'FF0000' }); c.border = { top: M, bottom: M, left: c.border.left, right: c.border.right }; }); });
    t.ws.views = [{ state: 'frozen', ySplit: 5 }];
    return t.ws;
  }
  function distributionSheet(wb, sheets) {
    // sheets: [{label, name, rows}] counting K4:K(3+rows)
    var cats = ['Weak', 'Acceptable', 'Good', 'Very Good', 'Outstanding'];
    var q = function (s) { return "'" + s.name.replace(/'/g, "''") + "'"; };
    var rows = cats.map(function (k) { return [k].concat(sheets.map(function (s) { return { formula: 'COUNTIF(' + q(s) + '!$K$4:$K$' + (3 + s.rows) + ',"' + k + '")' }; })); });
    var nc = sheets.length;
    var col = function (i) { return String.fromCharCode(67 + i); };
    rows.push(['Total'].concat(sheets.map(function (_, i) { return { formula: 'SUM(' + col(i) + '6:' + col(i) + '10)' }; })));
    rows.push(['Weak share'].concat(sheets.map(function (_, i) { return { formula: 'IF(' + col(i) + '11>0,' + col(i) + '6/' + col(i) + '11,0)' }; })));
    rows.push(['Outstanding share'].concat(sheets.map(function (_, i) { return { formula: 'IF(' + col(i) + '11>0,' + col(i) + '10/' + col(i) + '11,0)' }; })));
    var t = tableSheet(wb, 'Distribution Summary', 'How many of each result', 'Simulated options are not actual results.',
      ['Result'].concat(sheets.map(function (s) { return s.label; })), rows, [18].concat(sheets.map(function () { return 13; })),
      { fills: [PALE].concat(sheets.map(function (s) { return s.sim ? BLUE : CREAM; })), tab: '2E75B6' });
    [12, 13].forEach(function (r) { for (var i = 0; i < nc; i++) t.ws.getCell(r, 3 + i).numFmt = '0.0%'; });
    t.ws.getRow(11).eachCell(function (c) { c.font = font({ bold: true }); });
    return t.ws;
  }

  function templateBook(wb) {
    var ws = wb.addWorksheet('Progress'); setup(ws);
    [2.7, 13, 13, 13, 13, 13, 13, 13].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.mergeCells('B2:H2'); ws.getCell('B2').value = 'Subject name'; ws.getCell('B2').font = font({ size: 16, bold: true }); ws.getCell('B2').fill = fill(YEL); ws.getCell('B2').alignment = { horizontal: 'center' };
    var head = ['Phase', 'Group of Students', 'Grade/Year', 'No. of students in this group', '% Students assessed', 'students that made expected or better progress', 'students that made better than expected progress'];
    head.forEach(function (h, i) { var c = ws.getCell(3, 2 + i); c.value = h; c.font = font({ bold: true }); c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; c.border = { top: M, bottom: M, left: i === 0 ? M : T, right: i === head.length - 1 ? M : T }; c.fill = fill(i === 3 || i >= 5 ? YEL : 'FFFFFF'); });
    ws.getRow(3).height = 64;
    for (var g = 1; g <= 12; g++) {
      var r = 3 + g;
      [null, 'Full Cohort', 'G' + g, null, 100, null, null].forEach(function (v, i) {
        var c = ws.getCell(r, 2 + i); c.value = v; c.alignment = { horizontal: 'center' }; c.fill = fill(i === 3 || i >= 5 ? 'FFFFCC' : CREAM);
        c.border = { top: T, bottom: g === 12 ? M : T, left: i === 0 ? M : T, right: i === head.length - 1 ? M : T }; c.font = font({});
      });
    }
    var n = ws.getCell('B17'); n.value = 'Fill in the phase number and the three yellow columns for each grade. You can add a "Total" row after each phase; the page adds totals up for you either way.'; n.font = font({ italic: true, color: '595959' });
  }

  function build(kind, ctx) {
    // ctx: {parsed, model, sim, rules, sel}
    var wb = new ExcelJS.Workbook(); wb.creator = 'Progress Judgements'; wb.created = new Date();
    if (kind === 'template') { templateBook(wb); return wb; }
    var subj = (ctx.parsed && ctx.parsed.subject) || 'Progress';
    var r = ctx.rules, opts = (ctx.sim.options || []).slice(0, 6);
    var fixedSpec = C.sheetSpecModel(ctx.model, r, null, subj + ' (fixed)');
    var rawSpec = ctx.parsed && ctx.parsed.raw && ctx.parsed.raw.length ? C.sheetSpecRaw(ctx.parsed, subj) : null;
    function addOriginal() { if (rawSpec) specSheet(wb, 'Original (as received)', rawSpec, { tab: '808080' }); }
    function addFixed() { specSheet(wb, 'Fixed (real numbers)', fixedSpec, { formulas: true, pct: '0.0%', rules: r, banner: 'FIXED: real numbers. Totals add up and phase labels match the totals.', bannerColor: '548235', tab: '548235' }); }
    function addOption(i) {
      var o = opts[i];
      specSheet(wb, 'Option ' + C.letter(i) + ' (SIMULATED)', C.sheetSpecModel(o.model, r, ctx.model, subj + ' (SIMULATED, Option ' + C.letter(i) + ')'),
        { formulas: true, pct: '0.0%', rules: r, sim: true, tab: 'C00000', banner: 'SIMULATED, NOT ACTUAL RESULTS. Option ' + C.letter(i) + ': ' + o.students + ' more students make expected progress.' });
    }
    if (kind === 'fixed') { addOriginal(); addFixed(); correctionsSheet(wb, ctx.parsed, r); criteriaSheet(wb, r, C.counted(ctx.model.items, r.mode).length); }
    else if (kind === 'one') { addOption(ctx.sel); addFixed(); changesSheet(wb, ctx.sim, [ctx.sel]); }
    else {
      addOriginal(); criteriaSheet(wb, r, C.counted(ctx.model.items, r.mode).length); addFixed(); correctionsSheet(wb, ctx.parsed, r);
      var list = [];
      if (rawSpec) list.push({ label: 'In the file', name: 'Original (as received)', rows: rawSpec.body.length });
      list.push({ label: 'Fixed (real)', name: 'Fixed (real numbers)', rows: fixedSpec.body.length });
      opts.forEach(function (o, i) { list.push({ label: 'Option ' + C.letter(i) + ' (simulated)', name: 'Option ' + C.letter(i) + ' (SIMULATED)', rows: fixedSpec.body.length, sim: true }); });
      distributionSheet(wb, list);
      opts.forEach(function (_, i) { addOption(i); });
      if (opts.length) changesSheet(wb, ctx.sim, opts.map(function (_, i) { return i; }));
    }
    return wb;
  }
  return { build: build };
})(Core);
