/* Builds the output workbook with ExcelJS. The uploaded sheet is copied with its own formatting,
   then corrected (actual) and simulated versions get formulas, highlights and notes. */
var Book = (function (E) {
  var T = { style: 'thin', color: { argb: 'FF000000' } }, M = { style: 'medium', color: { argb: 'FF000000' } }, G = { style: 'thin', color: { argb: 'FFBFBFBF' } }, N = { style: 'medium', color: { argb: 'FF1F3864' } };
  var C = { yel: 'FFFFFF00', cream: 'FFFEF2CB', pale: 'FFFEFFB3', blue: 'FFDFE3FF', ink: 'FF0000B8', orange: 'FFFCE4D6', green: 'FF548235', red: 'FFC00000', grey: 'FF808080', white: 'FFFFFFFF', head: 'FF1F3864', chg: 'FFC6EFCE', chgInk: 'FF006100', chgBar: 'FF63BE7B' };
  function fill(a) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: a } }; }
  function font(o) { o = o || {}; var f = { name: 'Calibri', size: o.size || 11, bold: !!o.bold, italic: !!o.italic }; if (o.color) f.color = { argb: o.color }; return f; }
  var L = E.colLetter;
  var NAMES = {
    orig: 'Original (as received)', actual: 'Actual results (corrected)', comp: 'Comparison', counts: 'Judgement counts',
    totals: 'Total judgements', log: 'Corrections log', plan: 'Planning assumptions', settings: 'Settings',
    summary: 'Summary', notes: 'Assistant notes', criteria: 'Criteria', changes: 'Every simulated change'
  };
  function simSheetName(cfg) {
    var n = String(cfg.simTitle || 'Simulated scenario—not actual results').replace(/[\[\]:*?\/\\]/g, '-');
    return n.length <= 31 ? n : 'Simulated—not actual results';
  }
  function q(name) { return "'" + name.replace(/'/g, "''") + "'"; }
  function setup(ws) { ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 }; }

  function clone(dst, src, name, noNotes) {
    var ws = dst.addWorksheet(name);
    var model = JSON.parse(JSON.stringify(src.model));
    // Simulated sheets carry no cell comments (see changePanel): drop the ones copied from the school's sheet.
    if (noNotes) (model.rows || []).forEach(function (r) { (r.cells || []).forEach(function (c) { delete c.comment; }); });
    model.name = name; model.mergeCells = model.merges || [];
    delete model.id; model.tables = []; model.media = []; model.conditionalFormattings = model.conditionalFormattings || []; model.dataValidations = model.dataValidations || {};
    ws.model = Object.assign(model, { id: ws.id });
    ws.name = name;
    return ws;
  }

  /* ---------- Settings sheet (formulas point here) ---------- */
  var S = { acc: '$C$5', good: '$C$6', vg: '$C$7', out: '$C$8', vw: '$C$9', capW: '$C$10', capO: '$C$11', rule: '$C$12', ruleIdx: '$C$13', method: '$C$14', basis: '$C$15', vwAsWeak: '$C$16', assessed: '$C$17', subst: '$C$18' };
  var RULES = ['Not supplied (nearest point used for now)', 'Nearest whole point', 'Round down', 'Round up'];
  function ruleText(r) { return { unspecified: RULES[0], nearest: RULES[1], down: RULES[2], up: RULES[3] }[r] || RULES[0]; }
  function settingsSheet(wb, cfg) {
    var ws = wb.addWorksheet(NAMES.settings, { properties: { tabColor: { argb: 'FF2E75B6' } } }); setup(ws);
    [3, 44, 30, 60].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.getCell('B2').value = 'Settings used for this workbook'; ws.getCell('B2').font = font({ size: 14, bold: true });
    ws.getCell('B3').value = 'Every formula in the workbook reads these cells. Thresholds were not changed to reach the targets.'; ws.getCell('B3').font = font({ italic: true, color: 'FF595959' });
    var rows = [
      ['Acceptable: share making expected or better progress (at least)', cfg.acc, '0%'],
      ['Good: share making better than expected progress (at least)', cfg.good, '0%'],
      ['Very Good: share making better than expected progress (at least)', cfg.vg, '0%'],
      ['Outstanding: share making better than expected progress (at least)', cfg.out, '0%'],
      ['Very Weak: below this share making expected or better (blank = not used)', cfg.vw == null ? null : cfg.vw, '0%'],
      ['Most Weak judgements allowed', cfg.capWeak, '0%'],
      ['Most Outstanding judgements allowed', cfg.capOut, '0%'],
      ['Rounding rule for decimal averages', ruleText(cfg.rounding)],
      ['Rounding rule number (used by formulas)', { formula: 'MATCH(' + S.rule + ',{"' + RULES.join('","') + '"},0)' }],
      ['How a phase total is judged', cfg.totalMethod === 'points' ? 'Average of its grade judgements (points)' : 'From summed student counts'],
      ['Which judgements count toward the limits', { both: 'Grades and phase totals', grades: 'Grades only', phases: 'Phase totals only' }[cfg.basis]],
      ['Count Very Weak with Weak for the limit', cfg.vwCountsAsWeak ? 'Yes' : 'No'],
      ['Use the % assessed column', cfg.useAssessed ? 'Yes' : 'No'],
      ['Flag an adjustment as substantial above (percentage points)', cfg.substantialPts]
    ];
    (cfg.extraRules || []).forEach(function (r, i) { rows.push(['Extra rule ' + (i + 1), r.label || E.ruleLabel(r)]); });
    rows.forEach(function (r, i) {
      var rr = 5 + i, a = ws.getCell(rr, 2), b = ws.getCell(rr, 3);
      a.value = r[0]; b.value = r[1]; if (r[2]) b.numFmt = r[2];
      a.font = font(); b.font = font({ bold: true, color: 'FF0000FF' }); b.alignment = { horizontal: 'center' };
      a.fill = fill(C.cream); b.fill = fill(C.pale);
      [a, b].forEach(function (x, j) { x.border = { top: T, bottom: T, left: j === 0 ? M : T, right: j === 1 ? M : T }; });
    });
    ws.getCell(13, 3).font = font({ color: 'FF595959' });
    // points table, sorted ascending so MATCH(...,1) works
    var P0 = Math.max(21, 5 + rows.length + 2);
    ws.getCell(P0, 2).value = 'Points for each judgement'; ws.getCell(P0, 2).font = font({ size: 12, bold: true });
    var pts = E.CATS.filter(function (k) { return cfg.points[k] != null && cfg.points[k] !== ''; }).map(function (k) { return [k, +cfg.points[k]]; }).sort(function (a, b) { return a[1] - b[1]; });
    ['Judgement', 'Points'].forEach(function (h, i) { var x = ws.getCell(P0 + 1, 2 + i); x.value = h; x.font = font({ bold: true }); x.fill = fill(C.yel); x.border = { top: M, bottom: M, left: i === 0 ? M : T, right: i === 1 ? M : T }; x.alignment = { horizontal: 'center' }; });
    pts.forEach(function (p, i) {
      var a = ws.getCell(P0 + 2 + i, 2), b = ws.getCell(P0 + 2 + i, 3); a.value = p[0]; b.value = p[1];
      [a, b].forEach(function (x, j) { x.font = font({ bold: j === 1, color: j === 1 ? 'FF0000FF' : null }); x.fill = fill(j ? C.pale : C.cream); x.alignment = { horizontal: 'center' }; x.border = { top: T, bottom: i === pts.length - 1 ? M : T, left: j === 0 ? M : T, right: j === 1 ? M : T }; });
    });
    var last = P0 + 1 + pts.length;
    return { names: '$B$' + (P0 + 2) + ':$B$' + last, pts: '$C$' + (P0 + 2) + ':$C$' + last };
  }

  /* ---------- formulas for one copied table ---------- */
  function ru(t, a) { return 'ROUNDUP(' + t + '*' + a + '-0.000000001,0)'; }
  function judgeF(a, g, h) {
    var st = q(NAMES.settings) + '!';
    return 'IF(' + a + '<=0,"No data",IF(AND(' + st + S.vw + '<>"",' + g + '<' + ru(st + S.vw, a) + '),"Very Weak",IF(' + g + '<' + ru(st + S.acc, a) + ',"Weak",IF(' + h + '>=' + ru(st + S.out, a) + ',"Outstanding",IF(' + h + '>=' + ru(st + S.vg, a) + ',"Very Good",IF(' + h + '>=' + ru(st + S.good, a) + ',"Good","Acceptable"))))))';
  }
  function roundF(p) { var st = q(NAMES.settings) + '!'; return 'CHOOSE(' + st + S.ruleIdx + ',ROUND(' + p + ',0),ROUND(' + p + ',0),ROUNDDOWN(' + p + ',0),ROUNDUP(' + p + ',0))'; }

  // mode: 'actual' | 'sim'. model: built model (actual or simulated). base: actual model (for sim highlights)
  function writeTable(ws, ctx, mode, model, base, tabRef) {
    var cfg = ctx.cfg, parsed = ctx.parsed, st = q(NAMES.settings) + '!';
    var byId = {}; model.items.forEach(function (it) { byId[it.id] = it; });
    var baseById = {}; (base || model).items.forEach(function (it) { baseById[it.id] = it; });
    var fixes = E.findFixes(parsed, cfg).fixes;
    var helpers = {}, taken = [];
    // Tables side by side share rows, so each one's helper columns go to the right of everything before it.
    var sheetMax = 0; parsed.blocks.forEach(function (b) { Object.keys(b.col).forEach(function (k) { if (b.col[k] != null) sheetMax = Math.max(sheetMax, b.col[k]); }); });
    parsed.blocks.forEach(function (b) {
      var c = b.col, maxC = 0;
      Object.keys(c).forEach(function (k) { if (c[k] != null) maxC = Math.max(maxC, c[k]); });
      if (b.wide) maxC = Math.max(maxC, sheetMax);
      taken.forEach(function (t) { if (t.r1 <= b.endRow && b.headerRow <= t.r2) maxC = Math.max(maxC, t.c2); });
      var nx = maxC + 1, hA = nx++, hP = nx++, hK = c.K || nx++;
      var dG = c.G == null, dH = c.H == null, hG = dG ? nx++ : c.G, hH = dH ? nx++ : c.H;
      var hReal = null, hChg = null;
      // (the actual value and the change are shown in the changed cells and in the list under the table)
      var colOf = { E: c.E, G: hG, H: hH };
      helpers[b.id] = { A: hA, P: hP, K: hK, G: hG, H: hH };
      taken.push({ r1: b.headerRow, r2: b.endRow, c2: nx - 1 });
      // On simulated sheets the working columns are folded away (+ above the columns shows them), so the table reads like the school's own.
      if (mode === 'sim') [hA, hP, hK, hG, hH].forEach(function (col) { if (col > maxC) { ws.getColumn(col).outlineLevel = 1; ws.getColumn(col).hidden = true; } });
      var hr = b.headerRow;
      var heads = [[hA, 'Assessed students (calculated)'], [hP, 'Points (component) / average (total)']];
      if (!c.K) heads.push([hK, 'Judgement (calculated)']);
      if (dG) heads.push([hG, 'Expected or better, students (worked out from %)']);
      if (dH) heads.push([hH, 'Better than expected, students (worked out from %)']);
      heads.forEach(function (h) {
        var x = ws.getCell(hr, h[0]); x.value = (b.wide ? b.title + ': ' : '') + h[1]; x.font = font({ bold: true, color: 'FF595959', size: 9 }); x.fill = fill('FFF2F2F2');
        x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; x.border = { top: M, bottom: M, left: T, right: T };
        if (ws.getColumn(h[0]).width == null || ws.getColumn(h[0]).width < 12) ws.getColumn(h[0]).width = 14;
      });
      function cellRef(col, r) { return L(col) + r; }
      function setF(col, r, f, result, fmt) { var x = ws.getCell(r, col); x.value = { formula: f, result: result }; if (fmt) x.numFmt = fmt; return x; }
      function styleHelper(x, isT) { x.font = font({ bold: isT, color: 'FF595959' }); x.fill = fill('FFF7F7F7'); x.alignment = { horizontal: 'center', vertical: 'middle' }; x.border = { top: T, bottom: T, left: T, right: T }; }
      // A simulated change: bright green, bold, and the cell itself shows the actual value it replaces.
      function markChanged(x, kind, was, col) {
        var f0 = x.font || {}; x.fill = fill(C.chg); x.font = { name: f0.name || 'Calibri', size: f0.size || 11, bold: true, color: { argb: C.chgInk } };
        x.alignment = Object.assign({}, x.alignment || {}, { wrapText: false, shrinkToFit: false });
        var w = kind === 'p' ? (was == null ? '' : (Math.round(was * 1000) / 10).toFixed(1) + '%') : String(was == null ? '' : (typeof was === 'number' ? E.fmt(was) : was)).replace(/"/g, '');
        x.numFmt = kind === 'p' ? '0.0%" (was ' + w + ')"' : kind === 'j' ? 'General;General;General;@" (was ' + w + ')"' : '0" (was ' + w + ')"';
        var shown = kind === 'p' ? 6 + w.length + 7 : kind === 'j' ? String(x.value && x.value.result != null ? x.value.result : '').length + w.length + 7 : 4 + w.length + 7;
        var need = Math.ceil(shown * ((f0.size || 11) / 11) * 1.12) + 2;
        if (col && (ws.getColumn(col).width || 9) < need) ws.getColumn(col).width = need;
      }
      // grade rows
      b.rows.forEach(function (x0) {
        var id = b.id + ':' + x0.r, it = byId[id], bi = baseById[id]; if (!it) return;
        var r = x0.r;
        // edited or simulated counts
        var aRef = cellRef(hA, r), gRef = cellRef(hG, r), hRef = cellRef(hH, r);
        [['G', dG, c.I], ['H', dH, c.J]].forEach(function (q) {
          if (!q[1]) return;
          var k = q[0], pc = q[2] != null ? cellRef(q[2], r) : null, cell;
          var changed = mode === 'sim' && bi && it[k] !== bi[k];
          if (mode === 'actual' && pc && x0[k] == null) cell = setF(colOf[k], r, 'ROUND(IF(' + pc + '>1,' + pc + '/100,' + pc + ')*' + aRef + ',0)', it[k], '0');
          else { cell = ws.getCell(r, colOf[k]); cell.value = it[k]; cell.numFmt = '0'; }
          styleHelper(cell, false);
          if (changed) { markChanged(cell, 'n', bi[k], colOf[k]); cell.note = 'SIMULATED (planning assumption). Actual value ' + E.fmt(bi[k]) + ', change ' + (it[k] - bi[k] > 0 ? '+' : '') + E.fmt(it[k] - bi[k]) + '.'; }
          else if (mode === 'actual') cell.note = 'Worked out from the percentage x assessed students, rounded to whole students.';
        });
        ['E', 'G', 'H'].forEach(function (k) {
          if (c[k] == null) return;
          var cell = ws.getCell(r, c[k]), srcV = x0.orig ? x0.orig[k] : x0[k];
          if (it[k] !== srcV && it[k] != null) {
            cell.value = it[k];
            if (mode === 'sim' && bi && it[k] !== bi[k]) { markChanged(cell, 'n', bi[k], c[k]); cell.note = 'SIMULATED (planning assumption). Actual value ' + E.fmt(bi[k]) + ', change ' + (it[k] - bi[k] > 0 ? '+' : '') + E.fmt(it[k] - bi[k]) + '.'; }
            else if (bi && it[k] === bi[k] && it[k] !== srcV) { cell.note = 'Edited on the page. File value ' + E.fmt(srcV) + '.'; }
          }
        });
        var aF = (cfg.useAssessed && c.F != null) ? 'ROUND(' + cellRef(c.E, r) + '*IF(' + cellRef(c.F, r) + '>1,' + cellRef(c.F, r) + '/100,' + cellRef(c.F, r) + '),0)' : cellRef(c.E, r);
        styleHelper(setF(hA, r, aF, it.A, '0'), false);
        if (c.I != null && !(dG && mode === 'actual')) { var xi = setF(c.I, r, 'IF(' + aRef + '>0,' + gRef + '/' + aRef + ',0)', it.I || 0, '0.0%'); if (mode === 'sim' && bi && it.G !== bi.G) markChanged(xi, 'p', bi.I, c.I); }
        if (c.J != null && !(dH && mode === 'actual')) { var xj = setF(c.J, r, 'IF(' + aRef + '>0,' + hRef + '/' + aRef + ',0)', it.J || 0, '0.0%'); if (mode === 'sim' && bi && it.H !== bi.H) markChanged(xj, 'p', bi.J, c.J); }
        var xk = setF(hK, r, judgeF(aRef, gRef, hRef), it.judg);
        if (!c.K) styleHelper(xk, false);
        if (mode === 'sim' && bi && it.judg !== bi.judg) { markChanged(xk, 'j', bi.judg, hK); xk.note = 'SIMULATED. Actual judgement: ' + bi.judg + '.'; }
        else if (mode === 'actual' && x0.K && x0.K !== it.judg) { xk.fill = fill(C.orange); xk.note = 'Corrected. The file said ' + x0.K + '.'; }
        if (mode === 'actual') {
          if (!dG && c.I != null && x0.I != null && Math.abs(x0.I - (it.I || 0)) > 5e-5) { ws.getCell(r, c.I).fill = fill(C.orange); ws.getCell(r, c.I).note = 'Corrected. The file said ' + E.pct(x0.I, 2) + '.'; }
          if (!dH && c.J != null && x0.J != null && Math.abs(x0.J - (it.J || 0)) > 5e-5) { ws.getCell(r, c.J).fill = fill(C.orange); ws.getCell(r, c.J).note = 'Corrected. The file said ' + E.pct(x0.J, 2) + '.'; }
        }
        styleHelper(setF(hP, r, 'IFERROR(INDEX(' + st + ctx.pts.pts + ',MATCH(' + cellRef(hK, r) + ',' + st + ctx.pts.names + ',0)),"")', it.pts == null ? '' : it.pts, '0'), false);
      });
      // total rows that exist in the sheet
      b.totals.forEach(function (t0) {
        var r = t0.r, mem = t0.members.slice();
        var unit = model.units.filter(function (u) { return u.total && u.total.r === r; })[0]; if (!unit) return;
        var it = unit.total, bu = (base || model).units.filter(function (u) { return u.total && u.total.r === r; })[0], bi = bu ? bu.total : null;
        var contiguous = mem.length && mem[mem.length - 1] - mem[0] === mem.length - 1;
        function sumF(col) { return contiguous ? 'SUM(' + L(col) + mem[0] + ':' + L(col) + mem[mem.length - 1] + ')' : mem.map(function (m) { return L(col) + m; }).join('+'); }
        var fixed = {};
        ['E', 'G', 'H'].forEach(function (k) {
          if (colOf[k] == null) return;
          // Shared columns (students) belong to every table side by side, and a merged label must not be overwritten.
          if (k === 'E' && b.wide) return;
          var tgt = ws.getCell(r, colOf[k]); if (tgt.isMerged && tgt.master && tgt.master.address !== tgt.address) return;
          var x = setF(colOf[k], r, sumF(colOf[k]), it[k], '0');
          if (colOf[k] !== c[k]) styleHelper(x, true);
          if (mode === 'actual' && c[k] != null && t0[k] != null && Math.abs(t0[k] - it[k]) > 1e-6) { x.fill = fill(C.orange); x.note = 'Corrected. The file said ' + E.fmt(t0[k]) + ' (it should be the sum of the rows).'; fixed[k] = 1; }
          if (mode === 'sim' && bi && it[k] !== bi[k]) { markChanged(x, 'n', bi[k], colOf[k]); x.note = 'Changes because its rows changed. Actual ' + E.fmt(bi[k]) + '.'; }
        });
        var aRef = cellRef(hA, r);
        styleHelper(setF(hA, r, sumF(hA), it.A, '0'), true);
        ['I', 'J'].forEach(function (k) {
          if (c[k] == null) return;
          var cnt = k === 'I' ? hG : hH, v = k === 'I' ? it.I : it.J;
          var x = setF(c[k], r, 'IF(' + aRef + '>0,' + L(cnt) + r + '/' + aRef + ',0)', v || 0, '0.0%');
          if (mode === 'actual' && !(k === 'I' ? dG : dH) && t0[k] != null && Math.abs(t0[k] - (v || 0)) > 5e-5) { x.fill = fill(C.orange); x.note = 'Corrected. The file said ' + E.pct(t0[k], 2) + ' (a simple average, not total divided by total).'; }
          if (mode === 'sim' && bi && Math.abs((v || 0) - ((k === 'I' ? bi.I : bi.J) || 0)) > 1e-9) markChanged(x, 'p', k === 'I' ? bi.I : bi.J, c[k]);
        });
        // points average and total judgement
        var pRange = contiguous ? L(hP) + mem[0] + ':' + L(hP) + mem[mem.length - 1] : mem.map(function (m) { return L(hP) + m; }).join(',');
        var xp = setF(hP, r, 'IFERROR(AVERAGE(' + pRange + '),"")', it.avg == null ? '' : it.avg, '0.00'); styleHelper(xp, true);
        if (it.roundingFlag) { xp.fill = fill('FFFFF2CC'); xp.note = 'Average ' + it.avg.toFixed(2) + '. No rounding rule was supplied, so the nearest whole point is used for now. Please confirm the rule.'; }
        var kf = cfg.totalMethod === 'points'
          ? 'IFERROR(INDEX(' + st + ctx.pts.names + ',MATCH(' + roundF(L(hP) + r) + ',' + st + ctx.pts.pts + ',1)),"No data")'
          : judgeF(aRef, L(hG) + r, L(hH) + r);
        var xk = setF(hK, r, kf, it.judg);
        if (!c.K) styleHelper(xk, true);
        if (mode === 'actual' && t0.K && t0.K !== it.judg) { xk.fill = fill(C.orange); xk.note = 'Corrected. The file said ' + t0.K + '.'; }
        if (mode === 'sim' && bi && it.judg !== bi.judg) { markChanged(xk, 'j', bi.judg, hK); xk.note = 'SIMULATED. Actual judgement: ' + bi.judg + '. It changes because its grades changed.'; }
      });
      // phase label fix: re-merge the phase column per unit when labels were out of place
      if (c.phase != null && fixes.some(function (f) { return f.kind === 'label'; })) {
        var lo = b.headerRow + 1, hi = b.endRow;
        (ws.model.merges || []).slice().forEach(function (m) {
          var mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m); if (!mm) return;
          if (mm[1] === L(c.phase) && mm[3] === L(c.phase) && +mm[4] >= lo && +mm[2] <= hi) { try { ws.unMergeCells(m); } catch (e) { } }
        });
        var proto = ws.getCell(b.rows[0].r, c.phase).style;
        model.units.filter(function (u) { return u.blockId === b.id; }).forEach(function (u) {
          var rows = u.grades.map(function (g) { return g.r; }).concat(u.total && u.total.r ? [u.total.r] : []).sort(function (a, z) { return a - z; });
          if (!rows.length) return;
          for (var rr = rows[0]; rr <= rows[rows.length - 1]; rr++) { var cc = ws.getCell(rr, c.phase); cc.value = rr === rows[0] ? (isNaN(+u.phase) ? u.phase : +u.phase) : null; cc.style = JSON.parse(JSON.stringify(proto)); }
          if (rows[rows.length - 1] - rows[0] === rows.length - 1 && rows.length > 1) { try { ws.mergeCells(rows[0], c.phase, rows[rows.length - 1], c.phase); } catch (e) { } }
          if (mode === 'actual') { ws.getCell(rows[0], c.phase).note = 'Phase label moved so it matches the rows its total adds up.'; }
        });
      }
    });
    return helpers;
  }
  function banner(ws, ctx, text, color, help, tight) {
    var b = ctx.parsed.blocks[0], minC = Infinity, maxC = 0;
    ctx.parsed.blocks.forEach(function (bb) { Object.keys(bb.col).forEach(function (k) { if (bb.col[k] != null) { minC = Math.min(minC, bb.col[k]); maxC = Math.max(maxC, bb.col[k]); } }); });
    var hx = 0; Object.keys(help || {}).forEach(function (id) { Object.keys(help[id]).forEach(function (k) { if (help[id][k] > hx) hx = help[id][k]; }); });
    // Simulated sheets: exactly as wide as the table (their working columns are folded away).
    if (!tight) maxC = Math.max(maxC + 3, hx);
    var row1Free = true;
    for (var c = minC; c <= maxC; c++) { var v = ws.getCell(1, c).value; if (v != null && v !== '') row1Free = false; }
    if (row1Free && b.headerRow > 1) {
      (ws.model.merges || []).forEach(function (m) { if (/^[A-Z]+1:/.test(m)) { try { ws.unMergeCells(m); } catch (e) { } } });
      try { ws.mergeCells(1, minC, 1, maxC); } catch (e) { }
      var x = ws.getCell(1, minC); x.value = text; x.font = font({ size: 11, bold: true, color: C.white }); x.fill = fill(color); x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      var w = 0; for (var cw = minC; cw <= maxC; cw++) if (!ws.getColumn(cw).hidden) w += ws.getColumn(cw).width || 9;
      ws.getRow(1).height = text.length * 1.15 > w ? 34 : 22;
    } else {
      var y = ws.getCell(b.headerRow, maxC + 2); y.value = text; y.font = font({ bold: true, color: C.white }); y.fill = fill(color); ws.getColumn(maxC + 2).width = 48;
    }
  }

  /* ---------- the list of changes under a simulated table ---------- */
  function changePanel(ws, ctx, o, idx, help, actHelp) {
    var multi = ctx.parsed.blocks.length > 1, act = q(NAMES.actual) + '!', letter = String.fromCharCode(65 + idx);
    var minC = Infinity, last = 0;
    ctx.parsed.blocks.forEach(function (b) { Object.keys(b.col).forEach(function (k) { if (b.col[k] != null) minC = Math.min(minC, b.col[k]); }); last = Math.max(last, b.endRow); });
    ws.eachRow({ includeEmpty: false }, function (row, r) { row.eachCell({ includeEmpty: false }, function (cell, c) { if (!ws.getColumn(c).hidden && cell.value != null && cell.value !== '') last = Math.max(last, r); }); });
    var c0 = minC === Infinity ? 1 : minC, r = last + 3, NC = 11;
    var heads = [multi ? 'Subject' : 'Table', 'Grade', 'What changed', 'Actual (students)', 'Simulated (students)', 'Change (students)', 'Actual %', 'Simulated %', 'Change (% points)', 'Judgement: actual', 'Judgement: simulated'];
    // The list uses the visible columns from the table's first column on (folded working columns are skipped).
    var COLS = [], W = [17, 20, 30, 12, 12, 12, 11, 11, 12, 16, 18];
    for (var cc0 = c0; COLS.length < NC; cc0++) if (!ws.getColumn(cc0).hidden) COLS.push(cc0);
    COLS.forEach(function (cn, k) { var col = ws.getColumn(cn); if ((col.width || 9) < W[k]) col.width = W[k]; });
    function CL(k) { return L(COLS[k]); }
    function mergeRow(rr, a, z) { try { ws.mergeCells(rr, COLS[a], rr, COLS[z]); } catch (e) { } }
    function put(rr, cc, v, f) { var x = ws.getCell(rr, COLS[cc]); x.value = v; x.style = {}; x.border = { top: G, bottom: G, left: G, right: G }; x.alignment = { vertical: 'middle', horizontal: cc < 3 ? 'left' : 'center', wrapText: true, indent: cc < 3 ? 1 : 0 }; x.font = font({}); if (f) x.numFmt = f; ws.getRow(rr).height = Math.max(ws.getRow(rr).height || 0, 20); return x; }
    mergeRow(r, 0, NC - 1);
    var t = ws.getCell(r, c0); t.value = 'WHAT THE SIMULATION CHANGES: OPTION ' + letter + ' (SIMULATED, NOT ACTUAL RESULTS)'; t.font = font({ bold: true, size: 12, color: C.white }); t.fill = fill(C.red); t.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; ws.getRow(r).height = 22;
    r++;
    mergeRow(r, 0, NC - 1);
    var big = o.cells.slice().sort(function (a, z) { return (Math.abs(z.dG) + Math.abs(z.dH)) / (z.to.A || 1) - (Math.abs(a.dG) + Math.abs(a.dH)) / (a.to.A || 1); })[0];
    var sm = ws.getCell(r, c0);
    sm.value = o.changed.length + ' judgement' + (o.changed.length === 1 ? '' : 's') + ' change by moving ' + o.students + ' student' + (o.students === 1 ? '' : 's') + ' in ' + o.cells.length + ' grade' + (o.cells.length === 1 ? '' : 's') + ' (' + (Math.round(o.share * 1000) / 10).toFixed(1) + '% of assessed students). Largest change in one grade: ' + o.maxPts.toFixed(1) + ' percentage points' + (big ? ' (' + (multi ? big.to.block + ' ' : '') + big.to.grade + ')' : '') + (o.substantial ? ', which is a substantial adjustment.' : '.') + ' These are planning assumptions, not actual results.';
    sm.font = font({ size: 11 }); sm.alignment = { wrapText: true, vertical: 'top', horizontal: 'left', indent: 1 }; ws.getRow(r).height = 34;
    r += 2;
    heads.forEach(function (h, i) { var x = put(r, i, h); x.font = font({ bold: true, color: C.white, size: 10 }); x.fill = fill(C.head); x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
    ws.getRow(r).height = 30; var first = r + 1;
    function rowFor(x, k) {
      var hb = help[x.to.blockId], ha = actHelp[x.to.blockId]; if (!hb || !ha) return;
      r++; var rr = x.to.r, kk = k === 'G' ? 'I' : 'J';
      put(r, 0, x.to.block); put(r, 1, x.to.grade + (x.to.phase ? ' (' + E.phaseLabel(x.to.phase) + ')' : ''));
      put(r, 2, k === 'G' ? 'Expected progress or better' : 'Better than expected');
      put(r, 3, { formula: act + L(ha[k]) + rr, result: x.from[k] }, '0');
      var sv = put(r, 4, { formula: L(hb[k]) + rr, result: x.to[k] }, '0'); sv.fill = fill(C.chg); sv.font = font({ bold: true, color: C.chgInk });
      var ch = put(r, 5, { formula: CL(4) + r + '-' + CL(3) + r, result: x.to[k] - x.from[k] }, '+0;-0;0'); ch.font = font({ bold: true, color: C.chgInk, size: 12 });
      put(r, 6, { formula: 'IF(' + L(hb.A) + rr + '>0,' + CL(3) + r + '/' + L(hb.A) + rr + ',0)', result: x.from[kk] || 0 }, '0.0%');
      var sp = put(r, 7, { formula: 'IF(' + L(hb.A) + rr + '>0,' + CL(4) + r + '/' + L(hb.A) + rr + ',0)', result: x.to[kk] || 0 }, '0.0%'); sp.fill = fill(C.chg); sp.font = font({ bold: true, color: C.chgInk });
      put(r, 8, { formula: '(' + CL(7) + r + '-' + CL(6) + r + ')*100', result: ((x.to[kk] || 0) - (x.from[kk] || 0)) * 100 }, '+0.0;-0.0;0.0');
      var same = x.from.judg === x.to.judg;
      put(r, 9, { formula: act + L(ha.K) + rr, result: x.from.judg });
      var js = put(r, 10, { formula: L(hb.K) + rr, result: x.to.judg }); if (!same) { js.fill = fill(C.chg); js.font = font({ bold: true, color: C.chgInk }); }
    }
    o.cells.forEach(function (x) { if (x.dG) rowFor(x, 'G'); if (x.dH) rowFor(x, 'H'); });
    var lastChange = r;
    o.changed.filter(function (x) { return x.from.kind === 'total'; }).forEach(function (x) {
      var hb = help[x.to.blockId], ha = actHelp[x.to.blockId]; if (!hb || !ha || !x.to.r) return;
      r++; put(r, 0, x.to.block); put(r, 1, x.to.label || 'Total'); put(r, 2, 'Knock-on from its grades');
      for (var q2 = 3; q2 <= 8; q2++) { var e = put(r, q2, null); e.fill = fill('FFF2F2F2'); }
      put(r, 9, { formula: act + L(ha.K) + x.to.r, result: x.from.judg });
      var js = put(r, 10, { formula: L(hb.K) + x.to.r, result: x.to.judg }); js.fill = fill(C.chg); js.font = font({ bold: true, color: C.chgInk });
    });
    r++;
    var tl = put(r, 0, 'All changes'); tl.font = font({ bold: true });
    mergeRow(r, 0, 2);
    if (lastChange >= first) { var tt = put(r, 5, { formula: 'SUM(' + CL(5) + first + ':' + CL(5) + lastChange + ')', result: o.cells.reduce(function (a, x) { return a + x.dG + x.dH; }, 0) }, '+0;-0;0'); tt.font = font({ bold: true, color: C.chgInk, size: 12 }); }
    var jc = put(r, 10, o.changed.length + ' judgements'); jc.font = font({ bold: true });
    for (var q3 = 0; q3 < NC; q3++) ws.getCell(r, COLS[q3]).border = { top: N, bottom: N, left: G, right: G };
    // ExcelJS writes cell comments after the data bars, an order Excel rejects (it then drops the sheet).
    // Simulated cells already show "(was …)", so the comments are left out of simulated sheets.
    ws.eachRow({ includeEmpty: true }, function (row) { row.eachCell({ includeEmpty: true }, function (cell) {
      cell._comment = undefined;
      // ExcelJS also caches the comment in the cell's value model once the sheet model has been read.
      var vm = cell._value && cell._value.model; if (vm && vm.comment) delete vm.comment;
    }); });
    if (lastChange >= first) ws.addConditionalFormatting({ ref: CL(5) + first + ':' + CL(5) + lastChange, rules: [{ type: 'dataBar', priority: 1, gradient: false, cfvo: [{ type: 'num', value: 0 }, { type: 'max' }], color: { argb: C.chgBar } }] });
  }

  /* ---------- plain formatted tables ---------- */
  function table(ws, r0, c0, headers, rows, o) {
    o = o || {};
    // One house style for every table the workbook adds: navy headings, white rows, light grid lines.
    // Callers still say which cells are simulated (blue) or highlighted; those map to the same colours as the simulated sheets.
    function headFill(f) { return f === C.blue ? C.red : C.head; }
    function bodyFill(f) { return f === C.blue ? C.chg : (f === C.pale ? 'FFF2F2F2' : (f === C.cream || f === C.yel || !f ? C.white : f)); }
    function ink(c) { return c === C.ink ? C.chgInk : c; }
    headers.forEach(function (h, i) {
      var x = ws.getCell(r0, c0 + i); x.value = h; x.font = font({ bold: true, color: C.white }); x.fill = fill(headFill(o.headFills && o.headFills[i]));
      x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      x.border = { top: G, bottom: N, left: G, right: G };
    });
    ws.getRow(r0).height = o.headHeight || 32;
    var widths = o.widths ? o.widths.slice() : headers.map(function (h, i) {
      var m = Math.max.apply(null, String(h).split(/\s+/).map(function (w) { return w.length; }).concat([Math.min(String(h).length, 14)]));
      rows.forEach(function (row) { var v = row[i]; if (v == null) return; var t = typeof v === 'number' ? 8 : String(v).length; m = Math.max(m, Math.min(t, 48)); });
      return Math.max(10, m + 3);
    });
    rows.forEach(function (row, j) {
      var r = r0 + 1 + j, last = j === rows.length - 1;
      row.forEach(function (v, i) {
        var x = ws.getCell(r, c0 + i); x.value = v;
        var f = o.fills ? (typeof o.fills === 'function' ? o.fills(i, j, row) : o.fills[i]) : C.cream; x.fill = fill(bodyFill(f));
        x.font = font({ bold: o.bold && o.bold(i, j, row), color: ink(o.colors && o.colors(i, j, row)) });
        x.alignment = { horizontal: o.left && o.left[i] ? 'left' : 'center', vertical: 'middle', wrapText: true, indent: o.left && o.left[i] ? 1 : 0 };
        x.border = { top: G, bottom: last ? N : G, left: G, right: G };
        if (o.fmt && o.fmt[i]) x.numFmt = o.fmt[i];
      });
      ws.getRow(r).height = 20;
    });
    widths.forEach(function (w, i) { var col = ws.getColumn(c0 + i); if (o.widths || (col.width || 9) < w) col.width = w; });
    return r0 + rows.length;
  }
  function titled(wb, name, title, sub, tab) {
    var ws = wb.addWorksheet(name, { properties: { tabColor: { argb: tab } } }); setup(ws);
    ws.getColumn(1).width = 3;
    ws.getCell('B2').value = title; ws.getCell('B2').font = font({ size: 18, bold: true, color: C.head }); ws.getRow(2).height = 30;
    if (sub) { ws.getCell('B3').value = sub; ws.getCell('B3').font = font({ italic: true, color: 'FF595959' }); ws.getRow(3).height = 18; }
    return ws;
  }

  /* ---------- the whole workbook ---------- */
  function build(ctx) {
    // ctx: {srcWb, sheetName, parsed, cfg, base, sim, sel}
    var cfg = ctx.cfg, wb = new ExcelJS.Workbook(); wb.creator = 'Progress Judgements'; wb.created = new Date();
    wb.calcProperties = { fullCalcOnLoad: true };
    var src = ctx.srcWb.getWorksheet(ctx.sheetName) || ctx.srcWb.worksheets[0];
    var opt = ctx.sim && ctx.sim.options && ctx.sim.options[ctx.sel || 0];
    var simName = simSheetName(cfg);
    var origSrc = (ctx.origWb && ctx.origWb.getWorksheet(ctx.sheetName)) || src;
    var o = clone(wb, origSrc, NAMES.orig); o.properties.tabColor = { argb: C.grey };
    var a = clone(wb, src, NAMES.actual); a.properties.tabColor = { argb: C.green };
    var s = opt ? clone(wb, src, simName, true) : null; if (s) s.properties.tabColor = { argb: C.red };
    var pts = settingsSheet(wb, cfg); ctx.pts = pts;
    var help = writeTable(a, ctx, 'actual', ctx.base, null);
    banner(a, ctx, 'ACTUAL RESULTS, corrected. Orange cells were fixed (see Corrections log).', C.green, help);
    var SIM_KEY = ' Green = simulated change, showing the actual value it replaces. All changes are listed below.';
    if (s) { var helpS = writeTable(s, ctx, 'sim', opt.model, ctx.base, help); banner(s, ctx, String(cfg.simTitle).toUpperCase() + ' (Option ' + String.fromCharCode(65 + (ctx.sel || 0)) + ').' + SIM_KEY, C.red, helpS, true); changePanel(s, ctx, opt, ctx.sel || 0, helpS, help); }
    // Every other option gets its own sheet in the same layout, like the original audited workbook.
    var optNames = [];
    if (opt) (ctx.sim.options || []).slice(0, 6).forEach(function (o, i) {
      if (i === (ctx.sel || 0)) return;
      var nm = 'Option ' + String.fromCharCode(65 + i) + ' (SIMULATED)', w = clone(wb, src, nm, true); w.properties.tabColor = { argb: C.red };
      var hh = writeTable(w, ctx, 'sim', o.model, ctx.base, help);
      banner(w, ctx, 'SIMULATED, NOT ACTUAL RESULTS (Option ' + String.fromCharCode(65 + i) + ').' + SIM_KEY, C.red, hh, true);
      changePanel(w, ctx, o, i, hh, help);
      optNames.push(nm);
    });
    criteriaSheet(wb, ctx);
    if (opt) changesSheet(wb, ctx);
    comparison(wb, ctx, opt, simName, help);
    countsSheet(wb, ctx, opt);
    totalsSheet(wb, ctx, opt, simName, help);
    logSheet(wb, ctx);
    planSheet(wb, ctx, opt);
    if (ctx.summary) summarySheet(wb, ctx.summary);
    if (ctx.notes && (ctx.notes.transcript.length || ctx.notes.changes.length)) notesSheet(wb, ctx.notes);
    // order
    var order = (ctx.summary ? [NAMES.summary] : []).concat([NAMES.orig, NAMES.criteria, NAMES.actual, NAMES.log, NAMES.counts]).concat(s ? [simName] : []).concat(optNames).concat([NAMES.comp, NAMES.totals, NAMES.plan]).concat(opt ? [NAMES.changes] : []).concat(ctx.notes && wb.getWorksheet(NAMES.notes) ? [NAMES.notes] : []).concat([NAMES.settings]);
    order.forEach(function (n, i) { var w = wb.getWorksheet(n); if (w) w.orderNo = i; });
    wb.eachSheet(function (w) {
      setup(w);
      if (w.name === NAMES.orig) return;
      var school = w.name === NAMES.actual || w.name === simName || optNames.indexOf(w.name) >= 0;
      // Sheets the workbook makes read as pages: no grid lines. The school's own table keeps its look.
      if (!school) { var v = (w.views && w.views[0]) || {}; w.views = [Object.assign({}, v, { showGridLines: false, zoomScale: 100 })]; }
      fitRows(w, school ? lastTableRow(ctx) + 1 : 1);
    });
    return wb;
  }
  function lastTableRow(ctx) { var m = 0; ctx.parsed.blocks.forEach(function (b) { m = Math.max(m, b.endRow); }); return m; }
  // Excel does not resize rows to fit wrapped text when it opens a file, so every row with wrapped text gets a height that fits it.
  function fitRows(ws, fromRow) {
    var merged = {};
    Object.keys(ws._merges || {}).forEach(function (k) {
      var m = ws._merges[k].model || ws._merges[k], w = 0;
      for (var c = m.left; c <= m.right; c++) if (!ws.getColumn(c).hidden) w += ws.getColumn(c).width || 9;
      merged[m.top + ',' + m.left] = { w: w, tall: m.bottom > m.top };
    });
    ws.eachRow({ includeEmpty: false }, function (row, r) {
      if (r < fromRow) return;
      var need = 0;
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        if (cell.isMerged && cell.master && cell.master.address !== cell.address) return;
        var al = cell.alignment || {}; if (!al.wrapText || al.textRotation) return;
        // Only text wraps; numbers never do (Excel shows ### instead), so they never set the height.
        var v = cell.value, rv = v && typeof v === 'object' && !v.richText ? v.result : v;
        var txt = v && v.richText ? v.richText.map(function (t) { return t.text; }).join('') : (typeof rv === 'string' ? rv : '');
        if (!txt) return;
        var mg = merged[r + ',' + c]; if (mg && mg.tall) return;
        var w = mg ? mg.w : (ws.getColumn(c).width || 9), f = cell.font || {}, sz = f.size || 11;
        var per = (sz / 11) * (f.bold ? 1.06 : 1) * 0.9, cap = Math.max(1, (w - 1.2 - (al.indent || 0)) / per);
        var lines = txt.split('\n').reduce(function (n, l) { return n + Math.max(1, Math.ceil(l.length / cap)); }, 0);
        need = Math.max(need, lines * sz * 1.34 + 6);
      });
      if (need && need > (row.height || 15)) row.height = Math.min(409, Math.ceil(need));
    });
  }

  function comparison(wb, ctx, opt, simName, help) {
    var ws = titled(wb, NAMES.comp, 'Actual and simulated results, row by row', opt ? 'Every count, its simulated value, the difference and any judgement change. Simulated values are planning assumptions, not actual results.' : 'The actual results meet the targets, so there is no simulation.', 'FF2E75B6');
    var base = ctx.base, sm = opt ? opt.model : base, byId = {}; sm.items.forEach(function (it) { byId[it.id] = it; });
    var heads = ['Table', 'Group', 'Phase', 'Grade', 'Level', 'Students', 'Assessed', 'Expected or better: actual', 'Expected or better: simulated', 'Difference', 'Better than expected: actual', 'Better than expected: simulated', 'Difference', '% expected or better: actual', '% expected or better: simulated', '% better than expected: actual', '% better than expected: simulated', 'Judgement: actual', 'Judgement: simulated', 'Judgement changed?'];
    var r0 = 5, rows = [];
    base.items.forEach(function (b, i) {
      var s = byId[b.id] || b, r = r0 + 1 + i;
      rows.push([b.block, b.group, b.phase, b.kind === 'total' ? 'Total' : b.grade, b.kind === 'total' ? 'Phase total' : 'Grade', b.E, b.A, b.G, s.G,
        { formula: 'J' + r + '-I' + r, result: s.G - b.G }, b.H, s.H, { formula: 'M' + r + '-L' + r, result: s.H - b.H },
        { formula: 'IF(H' + r + '>0,I' + r + '/H' + r + ',0)', result: b.I || 0 }, { formula: 'IF(H' + r + '>0,J' + r + '/H' + r + ',0)', result: s.I || 0 },
        { formula: 'IF(H' + r + '>0,L' + r + '/H' + r + ',0)', result: b.J || 0 }, { formula: 'IF(H' + r + '>0,M' + r + '/H' + r + ',0)', result: s.J || 0 },
        refJ(NAMES.actual, b, ctx, help) || b.judg, opt ? (refJ(simName, s, ctx, help) || s.judg) : refJ(NAMES.actual, b, ctx, help) || b.judg,
        { formula: 'IF(S' + r + '=T' + r + ',"No","Yes")', result: s.judg === b.judg ? 'No' : 'Yes' }]);
    });
    var simCols = [8, 11, 14, 16, 18];
    table(ws, r0, 2, heads, rows, {
      headFills: heads.map(function (_, i) { return simCols.indexOf(i) >= 0 ? C.blue : C.yel; }),
      fills: function (i, j, row) { return simCols.indexOf(i) >= 0 ? C.blue : (row[4] === 'Phase total' ? C.pale : C.cream); },
      bold: function (i, j, row) { return row[4] === 'Phase total' || i === 9 || i === 12 || i === 19; },
      colors: function (i, j, row) { return simCols.indexOf(i) >= 0 ? C.ink : null; },
      fmt: [null, null, null, null, null, '0', '0', '0', '0', '+0;-0;0', '0', '0', '+0;-0;0', '0.0%', '0.0%', '0.0%', '0.0%'],
      headHeight: 46, widths: [14, 14, 8, 10, 11, 9, 9, 11, 11, 10, 11, 11, 10, 11, 11, 11, 11, 13, 13, 11]
    });
    ws.views = [{ state: 'frozen', ySplit: r0, xSplit: 5 }];
    // highlight judgement changes
    rows.forEach(function (row, j) { if (row[19].result === 'Yes') ['T', 'S'].forEach(function (L2) { ws.getCell(L2 + (r0 + 1 + j)).font = font({ bold: true, color: 'FFC00000' }); }); });
    ctx.compRows = { first: r0 + 1, last: r0 + rows.length };
  }
  function refJ(sheet, it, ctx, help) {
    if (!it.r) return null;
    var b = ctx.parsed.blocks.filter(function (x) { return x.id === it.blockId; })[0]; if (!b) return null;
    var col = help[b.id] ? help[b.id].K : b.col.K;
    return { formula: q(sheet) + '!' + L(col) + it.r, result: it.judg };
  }

  function countsSheet(wb, ctx, opt) {
    var cfg = ctx.cfg, ws = titled(wb, NAMES.counts, 'How many of each judgement', 'Grade-level and phase-level results are counted separately, then together. Checked against both limits in Settings.', 'FF2E75B6');
    var cr = ctx.compRows, cmp = q(NAMES.comp) + '!', st = q(NAMES.settings) + '!';
    var lev = cmp + '$F$' + cr.first + ':$F$' + cr.last, ja = cmp + '$S$' + cr.first + ':$S$' + cr.last, js = cmp + '$T$' + cr.first + ':$T$' + cr.last;
    var cats = E.CATS.slice().reverse();
    var groups = [['Grades', '"Grade"'], ['Phase totals', '"Phase total"'], ['All judgements', '"*"']];
    var heads = ['Judgement'];
    groups.forEach(function (g) { heads.push(g[0] + ': actual', g[0] + ': simulated'); });
    var r0 = 5, rows = cats.map(function (k) {
      var row = [k];
      groups.forEach(function (g) { row.push({ formula: 'COUNTIFS(' + lev + ',' + g[1] + ',' + ja + ',"' + k + '")' }, { formula: 'COUNTIFS(' + lev + ',' + g[1] + ',' + js + ',"' + k + '")' }); });
      return row;
    });
    var tr = r0 + cats.length + 1;
    rows.push(['Total'].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'SUM(' + col + (r0 + 1) + ':' + col + (tr - 1) + ')' }; })));
    var wkRows = cfg.vwCountsAsWeak ? [r0 + cats.indexOf('Weak') + 1, r0 + cats.indexOf('Very Weak') + 1] : [r0 + cats.indexOf('Weak') + 1];
    var oRow = r0 + cats.indexOf('Outstanding') + 1;
    rows.push(['Weak share' + (cfg.vwCountsAsWeak ? ' (incl. Very Weak)' : '')].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'IF(' + col + tr + '>0,(' + wkRows.map(function (x) { return col + x; }).join('+') + ')/' + col + tr + ',0)' }; })));
    rows.push(['Within the Weak limit?'].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'IF(' + col + (tr + 1) + '<=' + st + S.capW + '+0.0000001,"Yes","No")' }; })));
    rows.push(['Outstanding share'].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'IF(' + col + tr + '>0,' + col + oRow + '/' + col + tr + ',0)' }; })));
    rows.push(['Within the Outstanding limit?'].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'IF(' + col + (tr + 3) + '<=' + st + S.capO + '+0.0000001,"Yes","No")' }; })));
    var xr = cfg.extraRules || [];
    xr.forEach(function (rule, ri) {
      var catRows = (rule.cats || []).map(function (k) { return r0 + cats.indexOf(k) + 1; }).filter(function (x) { return x > r0; });
      var shareRow = tr + 5 + ri * 2;
      rows.push([(rule.label || E.ruleLabel(rule)) + ': share'].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'IF(' + col + tr + '>0,(' + (catRows.map(function (x) { return col + x; }).join('+') || '0') + ')/' + col + tr + ',0)' }; })));
      rows.push([(rule.label || E.ruleLabel(rule)) + ': met?'].concat(heads.slice(1).map(function (_, i) { var col = L(3 + i); return { formula: 'IF(' + col + shareRow + (rule.op === 'min' ? '>=' + rule.share + '-0.0000001' : '<=' + rule.share + '+0.0000001') + ',"Yes","No")' }; })));
    });
    // fill results so viewers without recalculation still see numbers
    var lists = {
      'Grades': [E.counted(ctx.base.items, 'grades'), E.counted((opt ? opt.model : ctx.base).items, 'grades')],
      'Phase totals': [E.counted(ctx.base.items, 'phases'), E.counted((opt ? opt.model : ctx.base).items, 'phases')],
      'All judgements': [E.counted(ctx.base.items, 'both'), E.counted((opt ? opt.model : ctx.base).items, 'both')]
    };
    var colLists = []; groups.forEach(function (g) { colLists.push(lists[g[0]][0], lists[g[0]][1]); });
    colLists.forEach(function (lst, i) {
      var k = E.check(lst, cfg), d = k.d;
      cats.forEach(function (cat, j) { rows[j][1 + i].result = d.counts[cat]; });
      rows[cats.length][1 + i].result = d.n;
      rows[cats.length + 1][1 + i].result = d.n ? k.weak / d.n : 0;
      rows[cats.length + 2][1 + i].result = k.weakOk ? 'Yes' : 'No';
      rows[cats.length + 3][1 + i].result = d.n ? k.out / d.n : 0;
      rows[cats.length + 4][1 + i].result = k.outOk ? 'Yes' : 'No';
      k.extra.forEach(function (x, ri) {
        var sh = 0; (xr[ri].cats || []).forEach(function (cat) { sh += d.counts[cat] || 0; });
        rows[cats.length + 5 + ri * 2][1 + i].result = d.n ? sh / d.n : 0;
        rows[cats.length + 6 + ri * 2][1 + i].result = (xr[ri].op === 'min' ? (d.n ? sh / d.n : 0) >= xr[ri].share - 1e-7 : (d.n ? sh / d.n : 0) <= xr[ri].share + 1e-7) ? 'Yes' : 'No';
      });
    });
    table(ws, r0, 2, heads, rows, {
      headFills: heads.map(function (_, i) { return i > 0 && i % 2 === 0 ? C.blue : C.yel; }),
      fills: function (i, j) { return i === 0 ? C.pale : (i % 2 === 0 ? C.blue : C.cream); },
      bold: function (i, j) { return j >= cats.length; },
      fmt: [null], widths: [30, 12, 12, 12, 12, 12, 12], headHeight: 40
    });
    var shareRows = [tr + 1, tr + 3], metRows = [tr + 2, tr + 4];
    xr.forEach(function (_, ri) { shareRows.push(tr + 5 + ri * 2); metRows.push(tr + 6 + ri * 2); });
    shareRows.forEach(function (r) { for (var i = 0; i < 6; i++) ws.getCell(r, 3 + i).numFmt = '0.0%'; });
    metRows.forEach(function (r) { for (var i = 0; i < 6; i++) { var x = ws.getCell(r, 3 + i); x.font = font({ bold: true, color: x.value.result === 'Yes' ? 'FF375623' : 'FFC00000' }); } });
    var note = ws.getCell(tr + 7 + xr.length * 2, 2);
    note.value = 'The targets are checked on: ' + { both: 'all judgements (grades and phase totals)', grades: 'grade judgements only', phases: 'phase total judgements only' }[cfg.basis] + '. With few judgements, a 20% limit allows very few Weak results (for example, 0 of 3).';
    note.font = font({ italic: true, color: 'FF595959' });
  }

  function totalsSheet(wb, ctx, opt, simName, help) {
    var cfg = ctx.cfg, ws = titled(wb, NAMES.totals, 'Total judgements from points', 'Each grade (component) has equal weight. Average = sum of points / number of components. Aggregate percentages come from summed counts and are separate from the points.', 'FF2E75B6');
    var heads = ['Table', 'Group', 'Phase', 'Components', 'Number', 'Points: actual', 'Average: actual', 'Rounded', 'Total judgement: actual', 'Points: simulated', 'Average: simulated', 'Rounded', 'Total judgement: simulated', 'Assessed (sum)', 'Expected or better (sum)', '% expected or better', 'Better than expected (sum)', '% better than expected', 'Rounding note'];
    var r0 = 5, rows = [], st = q(NAMES.settings) + '!';
    var sm = opt ? opt.model : ctx.base;
    ctx.base.units.forEach(function (u, i) {
      if (!u.total) return;
      var su = sm.units[i], t = u.total, ts = su.total, r = r0 + 1 + rows.length;
      var b = ctx.parsed.blocks.filter(function (x) { return x.id === u.blockId; })[0], hP = help[b.id] && help[b.id].P;
      function ptsF(sheet, unit) { var refs = unit.grades.filter(function (g) { return g.r; }).map(function (g) { return q(sheet) + '!' + L(hP) + g.r; }); return refs.length === unit.grades.length && hP ? 'SUM(' + refs.join(',') + ')' : null; }
      var pa = ptsF(NAMES.actual, u), ps = opt ? ptsF(simName, su) : pa;
      rows.push([u.block, u.group, u.phase, u.grades.map(function (g) { return g.grade + ' (' + g.judg + ')'; }).join(', '), u.grades.length,
        pa ? { formula: pa, result: t.sumPts } : t.sumPts, { formula: 'IF(F' + r + '>0,G' + r + '/F' + r + ',"")', result: t.avg },
        { formula: roundF('H' + r), result: t.rounded }, { formula: 'IFERROR(INDEX(' + st + ctx.pts.names + ',MATCH(I' + r + ',' + st + ctx.pts.pts + ',1)),"")', result: t.judgPoints },
        ps ? { formula: ps, result: ts.sumPts } : ts.sumPts, { formula: 'IF(F' + r + '>0,K' + r + '/F' + r + ',"")', result: ts.avg },
        { formula: roundF('L' + r), result: ts.rounded }, { formula: 'IFERROR(INDEX(' + st + ctx.pts.names + ',MATCH(M' + r + ',' + st + ctx.pts.pts + ',1)),"")', result: ts.judgPoints },
        t.A, t.G, { formula: 'IF(O' + r + '>0,P' + r + '/O' + r + ',0)', result: t.I || 0 }, t.H, { formula: 'IF(O' + r + '>0,R' + r + '/O' + r + ',0)', result: t.J || 0 },
        (t.roundingFlag || ts.roundingFlag) ? 'Decimal average and no rounding rule supplied: please confirm. Nearest whole point used for now.' : (cfg.rounding === 'unspecified' ? 'Whole-number average, no rounding needed.' : 'Rule: ' + ruleText(cfg.rounding) + '.')]);
    });
    if (!rows.length) rows.push(['No phase totals found', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    table(ws, r0, 2, heads, rows, {
      headFills: heads.map(function (_, i) { return i >= 9 && i <= 12 ? C.blue : C.yel; }),
      fills: function (i) { return i >= 9 && i <= 12 ? C.blue : (i === 18 ? 'FFFFF2CC' : C.cream); },
      bold: function (i) { return i === 8 || i === 12; }, colors: function (i) { return i >= 9 && i <= 12 ? C.ink : null; },
      left: { 3: true, 18: true }, fmt: [null, null, null, null, '0', '0', '0.00', '0', null, '0', '0.00', '0', null, '0', '0', '0.0%', '0', '0.0%'],
      widths: [14, 14, 8, 40, 9, 9, 10, 9, 14, 10, 10, 9, 14, 10, 11, 11, 11, 11, 40], headHeight: 46
    });
    ws.getCell(r0 + rows.length + 2, 2).value = 'Points: ' + E.CATS.slice().reverse().filter(function (k) { return cfg.points[k] != null && cfg.points[k] !== ''; }).map(function (k) { return k + ' = ' + cfg.points[k]; }).join(', ') + '. Rounding rule: ' + ruleText(cfg.rounding) + '.';
    ws.getCell(r0 + rows.length + 2, 2).font = font({ italic: true, color: 'FF595959' });
  }

  function logSheet(wb, ctx) {
    var f = E.findFixes(ctx.parsed, ctx.cfg);
    var ws = titled(wb, NAMES.log, 'Corrections made to the actual results', 'Only verifiable calculation and layout errors were corrected. No student counts were changed. Corrected cells are orange in "' + NAMES.actual + '".', 'FF548235');
    var rows = f.fixes.map(function (x, i) { return [i + 1, x.r ? 'Row ' + x.r : 'Whole table', x.what, x.detail, x.effect || 'None']; });
    if (!rows.length) rows = [[1, '', 'No calculation errors found', 'Every count, percentage and judgement already matched the rules.', 'None']];
    var last = table(ws, 5, 2, ['#', 'Where', 'What was wrong', 'Details', 'Effect on judgement'], rows, { left: { 2: true, 3: true }, widths: [5, 12, 32, 70, 18], fills: [C.cream, C.cream, C.pale, C.cream, C.cream] });
    ws.getCell(last + 2, 2).value = f.okRows + ' of ' + f.totalRows + ' grade rows were already correct.'; ws.getCell(last + 2, 2).font = font({ italic: true, color: 'FF375623' });
  }

  function planSheet(wb, ctx, opt) {
    var cfg = ctx.cfg, ws = titled(wb, NAMES.plan, 'Planning assumptions behind the simulated scenario', 'Hypothetical changes for internal analysis only. They are not findings, and no evidence is claimed for them.', 'FFC00000');
    if (!opt) { ws.getCell('B5').value = 'The corrected actual results already meet both limits, so no simulated changes were made.'; return; }
    var rows = opt.cells.map(function (c) {
      var b = c.from, s = c.to, parts = [];
      if (c.dG) parts.push((c.dG > 0 ? c.dG + ' more' : -c.dG + ' fewer') + ' of the ' + b.A + ' assessed students make expected or better progress (' + b.G + ' to ' + s.G + ')');
      if (c.dH) parts.push((c.dH > 0 ? c.dH + ' more' : -c.dH + ' fewer') + ' make better than expected progress (' + b.H + ' to ' + s.H + ')');
      return [s.block, s.group, s.phase, s.grade, (c.dG ? (c.dG > 0 ? '+' : '') + c.dG : '') + (c.dH ? ' / ' + (c.dH > 0 ? '+' : '') + c.dH : ''), c.pp, b.judg + ' to ' + s.judg, 'Planning assumption: ' + parts.join('; ') + '.', c.pp > cfg.substantialPts ? 'Substantial' : ''];
    });
    var last = table(ws, 5, 2, ['Table', 'Group', 'Phase', 'Grade', 'Students changed', 'Percentage points', 'Judgement', 'Assumption', 'Size'], rows,
      { left: { 7: true }, widths: [14, 14, 8, 10, 12, 12, 22, 70, 13], fmt: [null, null, null, null, null, '0.0'], fills: function (i) { return i === 4 || i === 5 ? C.blue : C.cream; },
        colors: function (i, j, row) { return i === 8 && row[8] ? 'FFC00000' : (i === 4 ? C.ink : null); }, bold: function (i) { return i === 8; } });
    var notes = [
      'Total: ' + opt.students + ' students across ' + opt.cells.length + ' grade' + (opt.cells.length === 1 ? '' : 's') + ' (' + (opt.share * 100).toFixed(1) + '% of all assessed students). Largest change in one grade: ' + opt.maxPts.toFixed(1) + ' percentage points.',
      opt.substantial ? 'At least one change is above ' + cfg.substantialPts + ' percentage points, which is a substantial adjustment.' : 'No single change is above ' + cfg.substantialPts + ' percentage points.',
      opt.changed.length + ' judgements change. All other judgements keep their actual values.',
      'Cohort sizes and assessed-student totals are unchanged. Counts are whole students within the assessed population, and better-than-expected never exceeds expected-or-better.',
      'Thresholds were not changed. Every judgement is recalculated from the simulated counts using the criteria in Settings.',
      'Each grade moves only as far as needed, so grades keep different results from each other.'
    ];
    var lim = [];
    if (cfg.simMaxStudents != null) lim.push('no grade gains more than ' + cfg.simMaxStudents + ' students');
    if (cfg.simMaxPts != null) lim.push('no grade moves more than ' + cfg.simMaxPts + ' percentage points');
    if ((cfg.simLocked || []).length) { var nm = {}; ctx.base.items.forEach(function (x) { nm[x.id] = x.grade; }); lim.push('these grades were left alone: ' + cfg.simLocked.map(function (id) { return nm[id] || id; }).join(', ')); }
    if (lim.length) notes.push('Limits set for this simulation: ' + lim.join('; ') + '.');
    if (Object.keys(cfg.simPinned || {}).length) notes.push('Some changes were chosen by hand rather than by the fewest-changes search; they are included above.');
    (opt.check.extra || []).forEach(function (x) { notes.push(x.label + ': ' + (x.ok ? 'met' : 'not met') + ' in this scenario (' + x.count + ' of ' + opt.check.n + ').'); });
    notes.forEach(function (t, i) { var x = ws.getCell(last + 2 + i, 2); x.value = '• ' + t; x.font = font({ color: i === 1 && opt.substantial ? 'FFC00000' : 'FF262626', bold: i === 0 }); });
  }

  /* ---------- criteria, every option's counts, every simulated change ---------- */
  function criteriaSheet(wb, ctx) {
    var cfg = ctx.cfg, ws = titled(wb, NAMES.criteria, 'How each result is decided', 'Thresholds mean “at least”. The numbers here come from the Settings sheet, so changing a rule there updates this sheet and every formula.', 'FF2E75B6');
    var st = q(NAMES.settings) + '!';
    var rows = [];
    if (cfg.vw != null) rows.push(['Very Weak', 'Students making expected or better progress', { formula: st + S.vw }, 'Fewer than this share make expected or better progress.']);
    rows.push(['Weak', 'Students making expected or better progress', { formula: st + S.acc }, 'Fewer than this share make expected or better progress.'],
      ['Acceptable', 'Students making expected or better progress', { formula: st + S.acc }, 'This share is reached, but fewer than the Good line did better than expected.'],
      ['Good', 'Students doing better than expected', { formula: st + S.good }, 'Acceptable, and at least this share did better than expected.'],
      ['Very Good', 'Students doing better than expected', { formula: st + S.vg }, 'Acceptable, and at least this share did better than expected.'],
      ['Outstanding', 'Students doing better than expected', { formula: st + S.out }, 'Acceptable, and at least this share did better than expected.']);
    rows.forEach(function (r) { r[2].result = { 'Very Weak': cfg.vw, 'Weak': cfg.acc, 'Acceptable': cfg.acc, 'Good': cfg.good, 'Very Good': cfg.vg, 'Outstanding': cfg.out }[r[0]]; });
    var last = table(ws, 5, 2, ['Result', 'What is checked', 'Line', 'Rule'], rows, { left: { 1: true, 3: true }, widths: [14, 40, 10, 62], fmt: [null, null, '0%'], fills: function (i) { return i === 0 ? C.pale : C.cream; } });
    var n = E.counted(ctx.base.items, cfg.basis).length, basisTxt = { both: 'grades and phase totals', grades: 'grades only', phases: 'phase totals only' }[cfg.basis];
    ws.getCell(last + 2, 2).value = 'Limits for the whole set of results'; ws.getCell(last + 2, 2).font = font({ bold: true, size: 12 });
    var lim = [['Weak', 'All ' + n + ' results (' + basisTxt + ')', { formula: st + S.capW, result: cfg.capWeak }, 'No more than this share may be Weak: at most ' + Math.floor(cfg.capWeak * n + 1e-9) + ' of ' + n + '.'],
      ['Outstanding', 'All ' + n + ' results (' + basisTxt + ')', { formula: st + S.capO, result: cfg.capOut }, 'No more than this share may be Outstanding: at most ' + Math.floor(cfg.capOut * n + 1e-9) + ' of ' + n + '.']];
    (cfg.extraRules || []).forEach(function (r) { lim.push(['Extra rule', 'All ' + n + ' results', r.share, r.label || E.ruleLabel(r)]); });
    var l2 = table(ws, last + 3, 2, ['Limit', 'Applies to', 'Maximum', 'Rule'], lim, { left: { 1: true, 3: true }, fmt: [null, null, '0%'], fills: function (i) { return i === 0 ? C.pale : C.cream; } });
    var nt = ws.getCell(l2 + 2, 2); nt.value = 'Phase totals: ' + (cfg.totalMethod === 'points' ? 'average of the grade judgements as points (' + E.CATS.filter(function (k) { return cfg.points[k] != null; }).slice().reverse().map(function (k) { return k + ' ' + cfg.points[k]; }).join(', ') + '), rounded ' + (cfg.rounding === 'unspecified' ? 'to the nearest point until a rule is confirmed' : cfg.rounding) + '.' : 'judged from their summed student counts.'); nt.font = font({ italic: true, color: 'FF595959' });
  }
  function changesSheet(wb, ctx) {
    var ws = titled(wb, NAMES.changes, 'Every simulated change, for every option (SIMULATED, not actual results)', 'Each option is a different way to meet the limits. Values are planning assumptions.', 'FFC00000');
    var multi = ctx.parsed.blocks.length > 1, rows = [], totals = [];
    (ctx.sim.options || []).slice(0, 6).forEach(function (o, i) {
      var L1 = 'Option ' + String.fromCharCode(65 + i) + (i === (ctx.sel || 0) ? ' (chosen)' : '');
      o.cells.forEach(function (x) { rows.push([L1, (multi ? x.to.block + ' ' : '') + x.to.grade, x.from.G, x.to.G, x.dG, x.from.judg, x.to.judg]); });
      o.changed.filter(function (x) { return x.from.kind === 'total'; }).forEach(function (x) { rows.push([L1, (multi ? x.to.block + ' ' : '') + x.to.label, '', '', '', x.from.judg, x.to.judg]); });
      rows.push([L1 + ' total', '', '', '', o.students, '', o.changed.length + ' results change']); totals.push(rows.length - 1);
      rows.push(['', '', '', '', '', '', '']);
    });
    rows.pop();
    table(ws, 5, 2, ['Option', 'Grade', 'Actual value', 'Simulated value', 'Change', 'Result before', 'Result after'], rows, {
      widths: [18, 26, 13, 15, 10, 15, 18], fmt: [null, null, '0', '0', '+0;-0;0'],
      fills: function (i, j, row) { return !row[0] ? C.white : (i === 3 || i === 4 || (i === 6 && row[5] !== row[6]) ? C.chg : C.cream); },
      colors: function (i, j, row) { return totals.indexOf(j) >= 0 ? 'FFC00000' : (i === 4 || i === 3 || i === 6 ? C.chgInk : null); }, bold: function (i, j) { return totals.indexOf(j) >= 0; }, left: { 0: true, 1: true } });
    ws.views = [{ state: 'frozen', ySplit: 5 }];
  }

  /* ---------- plain-language summary and assistant notes ---------- */
  // summary: {title, subtitle, sections: [{heading, text?, bullets?, table?: {head, rows}, tone?}]}
  function summarySheet(wb, sm) {
    var ws = wb.addWorksheet(NAMES.summary, { properties: { tabColor: { argb: 'FF262626' } } }); setup(ws);
    ws.getColumn(1).width = 3; [30, 16, 16, 16, 16, 18, 18].forEach(function (w, i) { ws.getColumn(2 + i).width = w; });
    var r = 2;
    var t = ws.getCell(r, 2); t.value = sm.title; t.font = font({ size: 20, bold: true, color: C.head }); ws.getRow(r).height = 32; r++;
    if (sm.subtitle) { var st0 = ws.getCell(r, 2); st0.value = sm.subtitle; st0.font = font({ italic: true, color: 'FF595959' }); r++; }
    r++;
    sm.sections.forEach(function (sec) {
      var h = ws.getCell(r, 2); h.value = sec.heading; h.font = font({ size: 14, bold: true, color: sec.tone === 'sim' ? 'FFC00000' : C.head }); ws.getRow(r).height = 24; h.alignment = { vertical: 'bottom' };
      for (var hc = 2; hc <= 8; hc++) ws.getCell(r, hc).border = { bottom: { style: 'thin', color: { argb: sec.tone === 'sim' ? 'FFC00000' : C.head } } };
      r++;
      if (sec.text) { ws.mergeCells(r, 2, r, 8); var x = ws.getCell(r, 2); x.value = sec.text; x.alignment = { wrapText: true, vertical: 'top' }; x.font = font(); r++; }
      (sec.bullets || []).forEach(function (b) { ws.mergeCells(r, 2, r, 8); var x = ws.getCell(r, 2); x.value = '\u2022 ' + b; x.alignment = { wrapText: true, vertical: 'top' }; x.font = font(); r++; });
      if (sec.table) { r = table(ws, r, 2, sec.table.head, sec.table.rows, { fills: function () { return sec.tone === 'sim' ? C.blue : C.cream; }, headFills: sec.table.head.map(function () { return sec.tone === 'sim' ? C.blue : C.yel; }), left: { 0: true }, headHeight: 30 }) + 1; }
      r++;
    });
  }
  function notesSheet(wb, nt) {
    var ws = wb.addWorksheet(NAMES.notes, { properties: { tabColor: { argb: 'FF0000DA' } } }); setup(ws);
    ws.getColumn(1).width = 3; ws.getColumn(2).width = 14; ws.getColumn(3).width = 100;
    ws.getCell('B2').value = 'What the assistant was asked, what it said, and what it changed'; ws.getCell('B2').font = font({ size: 14, bold: true });
    ws.getCell('B3').value = 'Changes listed here were made on the page before this workbook was built, and are already reflected in the other sheets.'; ws.getCell('B3').font = font({ italic: true, color: 'FF595959' });
    var r = 5;
    if (nt.changes.length) {
      ws.getCell(r, 2).value = 'Changes made'; ws.getCell(r, 2).font = font({ bold: true, size: 12 }); r++;
      nt.changes.forEach(function (c, i) { ws.getCell(r, 2).value = i + 1; ws.getCell(r, 3).value = c; ws.getCell(r, 3).alignment = { wrapText: true }; r++; });
      r++;
    }
    if (nt.transcript.length) {
      ws.getCell(r, 2).value = 'Conversation'; ws.getCell(r, 2).font = font({ bold: true, size: 12 }); r++;
      nt.transcript.forEach(function (m) {
        var a = ws.getCell(r, 2), b = ws.getCell(r, 3); a.value = m.role === 'user' ? 'Asked' : 'Assistant'; a.font = font({ bold: true, color: m.role === 'user' ? 'FF0000DA' : 'FF262626' }); a.alignment = { vertical: 'top' };
        var txt = String(m.text).replace(/\*\*/g, '').slice(0, 30000); b.value = txt; b.alignment = { wrapText: true, vertical: 'top' }; b.font = font();
        ws.getRow(r).height = Math.min(409, Math.max(16, 15 * txt.split('\n').reduce(function (n, l) { return n + Math.max(1, Math.ceil(l.length / 120)); }, 0)));
        r++;
      });
    }
  }

  /* ---------- a source workbook for pasted or typed data ---------- */
  function synthetic(rows, title, opt) {
    opt = opt || {};
    var pctMode = !!opt.pct;
    var wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Data');
    var head = pctMode
      ? ['Phase', 'Group of Students', 'Grade/Year', 'No. of students in this group', '% Students assessed', '% of students that made expected or better progress', '% of students that made better than expected progress', 'Judgement']
      : ['Phase', 'Group of Students', 'Grade/Year', 'No. of students in this group', '% Students assessed', 'students that made expected or better progress', 'students that made better than expected progress', '% of students that made expected or better progress', '% of students that made better than expected progress', 'Judgement'];
    var last = head.length - 1, key = pctMode ? [3, 5, 6] : [3, 5, 6];
    var widths = pctMode ? [3, 10, 14, 10, 13, 12, 16, 16, 13] : [3, 10, 14, 10, 13, 12, 14, 14, 14, 14, 13];
    widths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.mergeCells(2, 2, 2, 1 + head.length); var t = ws.getCell('B2'); t.value = title || 'Progress'; t.font = font({ size: 16, bold: true }); t.fill = fill(C.yel); t.alignment = { horizontal: 'center' };
    head.forEach(function (h, i) { var x = ws.getCell(3, 2 + i); x.value = h; x.font = font({ bold: true }); x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; x.border = { top: M, bottom: M, left: i === 0 ? M : T, right: i === last ? M : T }; x.fill = fill(i === last ? C.yel : C.white); });
    ws.getRow(3).height = 78;
    var r = 4;
    function frac(v) { return v == null || v === '' ? null : (+v > 1 ? +v / 100 : +v); }
    rows.forEach(function (x) {
      var vals = pctMode ? [x.phase, x.group || 'All students', x.grade, x.E, x.F == null ? 1 : frac(x.F), frac(x.I), frac(x.J), null]
        : [x.phase, x.group || 'All students', x.grade, x.E, x.F == null ? 100 : x.F, x.G, x.H, null, null, null];
      vals.forEach(function (v, i) {
        var c = ws.getCell(r, 2 + i); c.value = v === '' ? null : v; c.alignment = { horizontal: 'center' }; c.fill = fill(key.indexOf(i) >= 0 ? C.yel : C.cream); c.font = font({ bold: key.indexOf(i) >= 0 });
        if (pctMode && (i === 4 || i === 5 || i === 6)) c.numFmt = '0.0%';
        c.border = { top: T, bottom: T, left: i === 0 ? M : T, right: i === last ? M : T };
      });
      r++;
    });
    return wb;
  }
  function template() {
    var rows = []; for (var g = 1; g <= 12; g++) rows.push({ phase: null, group: 'All students', grade: 'G' + g, E: null, F: 100, G: null, H: null });
    var wb = synthetic(rows, 'Subject name');
    var ws = wb.getWorksheet('Data'); ws.name = 'Progress';
    ws.getCell('B17').value = 'Fill in the phase, the number of students and the two progress counts. Add a "Total" row after each phase if you like; the page adds them up either way.';
    ws.getCell('B17').font = font({ italic: true, color: 'FF595959' });
    return wb;
  }

  return { build: build, synthetic: synthetic, template: template, NAMES: NAMES, simSheetName: simSheetName };
})(Engine);
if (typeof module !== 'undefined') module.exports = Book;
