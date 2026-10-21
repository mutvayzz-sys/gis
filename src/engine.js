/* Progress Judgements engine: reads a sheet, applies the configured rules, finds corrections
   and searches for the smallest simulated changes. Pure functions, no DOM. */
var Engine = (function () {
  var CATS = ['Very Weak', 'Weak', 'Acceptable', 'Good', 'Very Good', 'Outstanding'];
  var ROLES = ['phase', 'group', 'grade', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
  var ROLE_LABEL = { phase: 'Phase', group: 'Group of students', grade: 'Grade or year', E: 'Number of students', F: '% students assessed', G: 'Made expected or better progress', H: 'Made better than expected progress', I: '% expected or better', J: '% better than expected', K: 'Judgement' };

  function defaults() {
    return {
      acc: 0.75, good: 0.50, vg: 0.61, out: 0.75, vw: null,
      points: { 'Outstanding': 6, 'Very Good': 5, 'Good': 4, 'Acceptable': 3, 'Weak': 2, 'Very Weak': 1 },
      totalMethod: 'points',          // 'points' | 'counts'
      rounding: 'unspecified',        // 'unspecified' | 'nearest' | 'down' | 'up'
      useAssessed: true,
      basis: 'both',                  // 'both' | 'grades' | 'phases'
      vwCountsAsWeak: true,
      capWeak: 0.20, capOut: 0.20,
      substantialPts: 10,
      simTitle: 'Simulated scenario—not actual results',
      extraRules: [],                 // [{label, cats:[...], op:'max'|'min', share:0..1}]
      simLocked: [],                  // grade ids the simulation may not change
      simMaxStudents: null,           // most students a simulation may add to one grade
      simMaxPts: null,                // most percentage points a simulation may move one grade
      simPinned: {},                  // the person's own changes {id: {G, H}} applied to the simulated scenario
      columns: {}                     // role -> column number override
    };
  }

  /* ---------- helpers ---------- */
  function needAt(A, t) { return Math.max(0, Math.ceil(t * A - 1e-9)); }
  function num(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).replace(/[,\s]/g, ''), pctSign = /%$/.test(s);
    var n = parseFloat(s.replace('%', ''));
    if (isNaN(n)) return null;
    return pctSign ? n / 100 : n;
  }
  function judge(A, G, H, c) {
    if (!(A > 0) || G == null || H == null) return 'No data';
    if (c.vw != null && G < needAt(A, c.vw)) return 'Very Weak';
    if (G < needAt(A, c.acc)) return 'Weak';
    if (H >= needAt(A, c.out)) return 'Outstanding';
    if (H >= needAt(A, c.vg)) return 'Very Good';
    if (H >= needAt(A, c.good)) return 'Good';
    return 'Acceptable';
  }
  function roundBy(avg, rule) {
    if (rule === 'down') return Math.floor(avg + 1e-9);
    if (rule === 'up') return Math.ceil(avg - 1e-9);
    return Math.floor(avg + 0.5 + 1e-9); // nearest, halves up (same as Excel ROUND for positive numbers)
  }
  function catFromPoints(p, c) {
    var best = null, bestPts = -Infinity, low = null, lowPts = Infinity;
    CATS.forEach(function (k) {
      var v = c.points[k]; if (v == null || v === '') return;
      if (v <= p + 1e-9 && v > bestPts) { best = k; bestPts = v; }
      if (v < lowPts) { low = k; lowPts = v; }
    });
    return best || low;
  }
  function ruleLabel(r) {
    var cats = r.cats || [], names;
    var orBetter = ['Outstanding', 'Very Good', 'Good', 'Acceptable'], i;
    for (i = 1; i <= orBetter.length; i++) if (cats.length === i && orBetter.slice(0, i).every(function (k) { return cats.indexOf(k) >= 0; })) { names = orBetter[i - 1] + (i > 1 ? ' or better' : ''); break; }
    if (!names) names = cats.join(' or ');
    return (r.op === 'min' ? 'At least ' : 'No more than ') + Math.round(r.share * 1000) / 10 + '% ' + names;
  }
  function phaseLabel(p) { p = String(p == null ? '' : p).trim(); return /^(phase|key stage|stage|cycle)\b/i.test(p) ? p : 'Phase ' + p; }
  function isWeakish(j, c) { return j === 'Weak' || (c.vwCountsAsWeak && j === 'Very Weak'); }

  /* ---------- reading a sheet ---------- */
  // grid: {cells: {"r,c": value}, masters: set of slave keys, maxR, maxC, hasFormulas}
  function gridFromAoa(aoa) {
    var cells = {}, maxR = 0, maxC = 0;
    aoa.forEach(function (row, r) { (row || []).forEach(function (v, c) { if (v != null && v !== '') { cells[(r + 1) + ',' + (c + 1)] = v; maxR = Math.max(maxR, r + 1); maxC = Math.max(maxC, c + 1); } }); });
    return { cells: cells, maxR: maxR, maxC: maxC, hasFormulas: undefined };
  }
  function get(grid, r, c) { var v = grid.cells[r + ',' + c]; return v === undefined ? null : v; }
  function text(v) { return v == null ? '' : String(v).trim(); }

  var EXP = /expected or better|expected or above|at or above|met or exceeded|met\s*(?:&|and|\/|\+)\s*exceed|met (?:the )?(?:projected|expected|growth|target|typical)|achieved (?:expected|growth|target|typical)|on track|at expected|meeting (?:expected|growth|target)/;
  var BET = /better than expected|above expected|exceeded (?:the )?(?:projected|expected|growth|target|typical)|^exceeded|above (?:target|typical)|accelerated|high growth|^above$|beyond expected|more than expected/;
  function detectHeader(grid, r, loose) {
    var col = {}, hits = 0;
    for (var c = 1; c <= grid.maxC; c++) {
      var s = text(get(grid, r, c)).toLowerCase().replace(/\s+/g, ' ');
      if (!s || s.length > 140) continue;
      var pct = /%|percent|\brate\b|proportion/.test(s);
      if (!pct && /phase|key stage|cycle|^stage$|school section/.test(s) && col.phase == null) col.phase = c;
      else if (!pct && col.E == null && /(?:no\.?|number|num|count|total|^n)\s*(?:of\s*)?(?:students|pupils|learners|tested|assessed|children)|students in|cohort size|^n$|^students$|^pupils$|^learners$|^(?:students |pupils )?tested$|enrol/.test(s)) { col.E = c; hits++; }
      else if (/assessed|participation|took the test/.test(s) && !/^(?:no\.?|number)/.test(s) && col.F == null) col.F = c;
      else if (!pct && /grade|year|\byr\b|class|level/.test(s) && !EXP.test(s) && !BET.test(s) && col.grade == null) { col.grade = c; hits++; }
      else if (/group|subgroup|category|cohort|subject|course/.test(s) && !EXP.test(s) && !BET.test(s) && col.group == null) col.group = c;
      else if (EXP.test(s)) { if (pct) { if (col.I == null) { col.I = c; hits++; } } else if (col.G == null) { col.G = c; hits++; } }
      else if (BET.test(s)) { if (pct) { if (col.J == null) { col.J = c; hits++; } } else if (col.H == null) { col.H = c; hits++; } }
      else if (/judg|rating|evaluation|outcome/.test(s) && col.K == null) col.K = c;
    }
    if (col.G != null && col.I != null) hits--; if (col.H != null && col.J != null) hits--;
    if (loose) return col;
    return hits >= 3 && col.grade != null && (col.G != null || col.I != null) ? col : null;
  }

  var JUDG_WORD = /^(very weak|weak|acceptable|good|very good|outstanding)$/i;
  // Finds subjects laid out side by side: the same measure heading repeated across the header row,
  // usually with subject names in the row above. Returns one column map per subject, or null.
  // A heading may carry the subject in front of the measure ("Reading % met or exceeded").
  function measureSplit(raw) {
    var t = text(raw).replace(/\s+/g, ' '), low = t.toLowerCase();
    var m = EXP.exec(low) || BET.exec(low) || /judg|rating|evaluation/.exec(low);
    if (!m) return null;
    var prefix = t.slice(0, m.index).replace(/[\s%:\-–(]+$/, '').trim();
    var key = low.slice(m.index).trim(), pct = /%|percent|rate/.test(low);
    return { prefix: prefix, key: key, pct: pct };
  }
  function splitWide(grid, r, end, col) {
    var heads = {}, starts = [], prefixed = {};
    for (var c = 1; c <= grid.maxC; c++) {
      var t = text(get(grid, r, c)).toLowerCase().replace(/\s+/g, ' ');
      if (t && (EXP.test(t) || BET.test(t))) (heads[t] = heads[t] || []).push(c);
      var ms = measureSplit(get(grid, r, c));
      // A subject name is a short label, not words like "students that made" or "% of".
      var okPrefix = ms && ms.prefix && ms.prefix.length <= 40 && /[a-z]/i.test(ms.prefix) && !/%|percent|\b(students?|pupils?|learners?|children|made|making|that|who|which|no\.?|number|of|the|share|proportion|rate|total)\b/i.test(ms.prefix);
      if (okPrefix && (EXP.test(ms.key) || BET.test(ms.key))) (prefixed[ms.key] = prefixed[ms.key] || []).push({ c: c, prefix: ms.prefix });
    }
    Object.keys(heads).forEach(function (t) { if (heads[t].length > starts.length) starts = heads[t]; });
    // Subject names inside the headings: the same measure under different subject prefixes.
    var subjTitles = null;
    Object.keys(prefixed).forEach(function (k) {
      var list = prefixed[k], names = list.map(function (x) { return x.prefix.toLowerCase(); });
      if (list.length > starts.length && list.length >= 2 && new Set(names).size === list.length) { starts = list.map(function (x) { return x.c; }); subjTitles = list.map(function (x) { return x.prefix; }); }
    });
    if (starts.length < 2) return null;
    var shared = {}; ['phase', 'group', 'grade', 'E', 'F'].forEach(function (k) { if (col[k] != null && col[k] < starts[0]) shared[k] = col[k]; });
    var width = starts[1] - starts[0];
    var lastRow = Math.min(end, r + 60);
    return starts.map(function (st0, gi) {
      var stop = gi + 1 < starts.length ? starts[gi + 1] - 1 : Math.min(grid.maxC, st0 + width - 1);
      var title = subjTitles ? subjTitles[gi] : null;
      for (var tc = st0; tc <= stop && !title; tc++) { var tv = text(get(grid, r - 1, tc)); if (tv && tv.length < 80) title = tv; }
      // A subject's columns end where the next subject's prefix begins.
      if (subjTitles) { var lowT = subjTitles[gi].toLowerCase(); for (var ec = st0 + 1; ec <= stop; ec++) { var ht = text(get(grid, r, ec)).toLowerCase(); if (ht && ht.indexOf(lowT) !== 0 && subjTitles.some(function (x) { return ht.indexOf(x.toLowerCase()) === 0; })) { stop = ec - 1; break; } } }
      var cm = Object.assign({}, shared), numeric = [];
      // Own headings for each column win; otherwise go by what the data looks like.
      var own = 0; for (var cc = st0; cc <= stop; cc++) if (text(get(grid, r, cc))) own++;
      for (cc = st0; cc <= stop; cc++) {
        var h = text(get(grid, r, cc)).toLowerCase(), pct = /%|percent|rate/.test(h);
        if (title && h.indexOf(title.toLowerCase()) === 0) h = h.slice(title.length);
        h = h.replace(/^[\s%:\-–(]+/, '').replace(/^(?:of\s+)?(?:students|pupils)\s+(?:who\s+)?/, '');
        if (own > 1 && h) {
          if (/judg|rating|evaluation/.test(h)) { cm.K = cc; continue; }
          if (BET.test(h)) { cm[pct ? 'J' : 'H'] = cc; continue; }
          if (EXP.test(h)) { cm[pct ? 'I' : 'G'] = cc; continue; }
        }
        var nums = 0, words = 0, small = 0, total = 0;
        for (var rr = r + 1; rr <= lastRow; rr++) { var v = get(grid, rr, cc); if (v == null || v === '') continue; total++; if (typeof v === 'number') { nums++; if (v <= 1) small++; } else if (JUDG_WORD.test(text(v))) words++; else if (!isNaN(parseFloat(String(v).replace('%', '')))) { nums++; } }
        if (!total) { if (own <= 1 && numeric.length < 2 && cm.K == null) numeric.push({ c: cc, pct: true }); continue; }
        if (words >= nums && words) cm.K = cm.K || cc;
        else if (nums) numeric.push({ c: cc, pct: small === nums || pct });
      }
      // Numbers without a recognised heading fill whichever measure is still missing, expected first.
      numeric.forEach(function (n) {
        if (cm.G == null && cm.I == null) cm[n.pct ? 'I' : 'G'] = n.c;
        else if (cm.H == null && cm.J == null) cm[n.pct ? 'J' : 'H'] = n.c;
      });
      return { col: cm, title: title };
    }).filter(function (p) { return p.col.G != null || p.col.I != null; });
  }
  function parse(grid, opt) {
    opt = opt || {};
    var over = opt.columns || {};
    var heads = [];
    if (opt.headerRow) heads.push({ r: +opt.headerRow, col: detectHeader(grid, +opt.headerRow, true) });
    else for (var r = 1; r <= grid.maxR; r++) { var h = detectHeader(grid, r); if (h) heads.push({ r: r, col: h }); }
    var blocks = [], nb = 0;
    heads.forEach(function (h, i) {
      var end = i + 1 < heads.length ? heads[i + 1].r - 1 : grid.maxR;
      var prevEnd = i > 0 ? heads[i - 1].r : 0;
      var title = null;
      for (var tr = h.r - 1; tr > prevEnd && !title; tr--) for (var tc = 1; tc <= grid.maxC && !title; tc++) { var tv = get(grid, tr, tc); if (typeof tv === 'string' && tv.trim() && tv.length < 80) title = tv.trim(); }
      // Several subjects side by side under one header row become one table each.
      var parts = splitWide(grid, h.r, end, h.col) || [{ col: h.col, title: null }];
      parts.forEach(function (pt) {
        var col = Object.assign({}, pt.col);
        ROLES.forEach(function (k) { if (over[k] !== undefined) col[k] = over[k] === '' ? null : over[k]; });
        blocks.push({ id: 'b' + (nb++), headerRow: h.r, endRow: end, col: col, title: pt.title || title || (opt.sheetName || 'Table ' + (i + 1)), sheetTitle: title, wide: parts.length > 1, rows: [], totals: [], skipped: [] });
      });
    });
    blocks.forEach(function (b) {
      var c = b.col, carryP = null, carryG = null, pend = [], pendExplicit = null;
      for (var r = b.headerRow + 1; r <= b.endRow; r++) {
        var gv = c.grade != null ? text(get(grid, r, c.grade)) : '';
        if (!gv) {
          // A total row may carry its label in the phase or group column, or in a merged cell to the left.
          var firstData = Math.min.apply(null, ['E', 'F', 'G', 'H', 'I', 'J'].map(function (k) { return c[k] != null ? c[k] : 999; }));
          for (var lc = 1; lc < Math.min(firstData, 30) && !gv; lc++) { var lt = text(get(grid, r, lc)); if (/^(total|overall|all grades|phase total)\b/i.test(lt) || /\btotal\b/i.test(lt)) gv = lt; }
          if (!gv) continue;
        } else if (/^\d+$/.test(gv)) gv = 'Grade ' + gv;
        var pl = c.phase != null ? get(grid, r, c.phase) : null;
        if (pl != null && text(pl) !== '') { carryP = text(pl); if (pendExplicit == null) pendExplicit = carryP; }
        var gl = c.group != null ? get(grid, r, c.group) : null;
        if (gl != null && text(gl) !== '') carryG = text(gl);
        var rec = {
          r: r, grade: gv, label: carryP, group: carryG || 'All students',
          E: num(get(grid, r, c.E)), F: c.F != null ? num(get(grid, r, c.F)) : null,
          G: num(get(grid, r, c.G)), H: num(get(grid, r, c.H)),
          I: c.I != null ? num(get(grid, r, c.I)) : null, J: c.J != null ? num(get(grid, r, c.J)) : null,
          K: c.K != null ? text(get(grid, r, c.K)) || null : null
        };
        rec.orig = { E: rec.E, G: rec.G, H: rec.H };
        if (/^(total|overall|all grades|phase total)/i.test(gv) || /\btotal\b/i.test(gv)) {
          var ph = pendExplicit != null ? pendExplicit : (pend[0] ? pend[0].label : String(b.totals.length + 1));
          pend.forEach(function (x) { x.phase = ph; });
          rec.phase = ph; rec.members = pend.map(function (x) { return x.r; });
          b.totals.push(rec); pend = []; pendExplicit = null;
        } else {
          if (rec.G == null && rec.H == null && rec.I == null && rec.J == null) { if (rec.E != null) b.skipped.push(rec.grade); continue; }
          b.rows.push(rec); pend.push(rec);
        }
      }
      pend.forEach(function (x) { x.phase = x.label != null ? x.label : '1'; });
      // "Assessed" may be a percentage (100), a fraction (1) or a count of students; store it as a percentage.
      if (c.F != null) {
        var fr = b.rows.filter(function (x) { return x.F != null && x.E; });
        var asCount = fr.length && !fr.some(function (x) { return x.F > x.E; }) && !fr.every(function (x) { return x.F <= 1; }) && !fr.every(function (x) { return x.F === 100; });
        if (asCount) b.rows.forEach(function (x) { if (x.F != null && x.E) x.F = x.F / x.E * 100; });
      }
    });
    blocks = blocks.filter(function (b) { return b.rows.length; });
    return { blocks: blocks, hasFormulas: grid.hasFormulas, sheetName: opt.sheetName, source: opt.source, detected: heads.map(function (h) { return h.col; }) };
  }
  // Header names for the column picker
  function headerOptions(grid, parsed) {
    var out = [];
    if (!parsed || !parsed.blocks.length) return out;
    var hr = parsed.blocks[0].headerRow;
    for (var c = 1; c <= grid.maxC; c++) { var t = text(get(grid, hr, c)); if (t) out.push({ col: c, label: colLetter(c) + ': ' + t.replace(/\s+/g, ' ').slice(0, 60) }); }
    return out;
  }
  function colLetter(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }

  /* ---------- model ---------- */
  function assessedOf(x, c) {
    if (!c.useAssessed || x.F == null) return x.E;
    var f = x.F > 1 ? x.F / 100 : x.F;
    return Math.round((x.E || 0) * f);
  }
  function build(parsed, c, override) {
    // override: {rowKey: {G, H}} simulated counts
    override = override || {};
    var units = [], items = [];
    parsed.blocks.forEach(function (b) {
      var map = {}, order = [];
      b.rows.forEach(function (x) {
        var o = override[b.id + ':' + x.r] || {};
        var ph = o.phase != null && o.phase !== '' ? String(o.phase) : x.phase;
        var key = b.id + '|' + x.group + '|' + ph;
        if (!map[key]) { map[key] = { key: key, block: b.title, blockId: b.id, group: x.group, phase: ph, grades: [], total: null }; order.push(key); }
        var A0 = assessedOf(x, c), frac = function (v) { return v == null ? null : (v > 1 ? v / 100 : v); };
        var G0 = x.G != null ? x.G : (x.I != null && A0 != null ? Math.round(frac(x.I) * A0) : null);
        var H0 = x.H != null ? x.H : (x.J != null && A0 != null ? Math.round(frac(x.J) * A0) : null);
        var it = { id: b.id + ':' + x.r, kind: 'grade', blockId: b.id, r: x.r, block: b.title, group: x.group, phase: ph, phaseSet: ph !== x.phase, grade: x.grade, label: x.grade,
          E: x.E, A: A0, G: o.G != null ? o.G : G0, H: o.H != null ? o.H : H0, derived: x.G == null || x.H == null };
        it.I = it.A > 0 ? it.G / it.A : null; it.J = it.A > 0 ? it.H / it.A : null;
        it.judg = judge(it.A, it.G, it.H, c); it.pts = c.points[it.judg] != null ? +c.points[it.judg] : null;
        map[key].grades.push(it);
      });
      order.forEach(function (key) {
        var u = map[key];
        var src = b.totals.filter(function (t) { return t.phase === u.phase && u.grades.some(function (g) { return t.members.indexOf(g.r) >= 0; }); })[0] || null;
        if (u.grades.length >= 2 || src) {
          var t = { id: b.id + ':T:' + u.group + ':' + u.phase, kind: 'total', blockId: b.id, r: src ? src.r : null, block: b.title, group: u.group, phase: u.phase, grade: 'Total', label: phaseLabel(u.phase) + ' total', E: 0, A: 0, G: 0, H: 0, src: src };
          u.grades.forEach(function (g) { t.E += g.E || 0; t.A += g.A || 0; t.G += g.G || 0; t.H += g.H || 0; });
          t.I = t.A > 0 ? t.G / t.A : null; t.J = t.A > 0 ? t.H / t.A : null;
          t.judgCounts = judge(t.A, t.G, t.H, c);
          var pts = u.grades.map(function (g) { return g.pts; }).filter(function (p) { return p != null; });
          t.sumPts = pts.reduce(function (a, b2) { return a + b2; }, 0); t.nPts = pts.length;
          t.avg = pts.length ? t.sumPts / pts.length : null;
          t.rounded = t.avg != null ? roundBy(t.avg, c.rounding) : null;
          t.judgPoints = t.avg != null ? catFromPoints(t.rounded, c) : 'No data';
          t.roundingFlag = c.totalMethod === 'points' && c.rounding === 'unspecified' && t.avg != null && Math.abs(t.avg - Math.round(t.avg)) > 1e-9;
          t.judg = c.totalMethod === 'points' ? t.judgPoints : t.judgCounts;
          t.pts = c.points[t.judg] != null ? +c.points[t.judg] : null;
          u.total = t;
        }
        units.push(u);
      });
    });
    units.forEach(function (u) { u.grades.forEach(function (g) { items.push(g); }); if (u.total) items.push(u.total); });
    return { units: units, items: items };
  }
  function counted(items, basis) {
    return items.filter(function (it) { return basis === 'both' || (basis === 'grades' ? it.kind === 'grade' : it.kind === 'total'); });
  }
  function dist(list) {
    var counts = {}; CATS.forEach(function (k) { counts[k] = 0; });
    list.forEach(function (it) { if (counts[it.judg] != null) counts[it.judg]++; });
    return { counts: counts, n: list.length };
  }
  function check(list, c) {
    var d = dist(list), n = d.n;
    var w = d.counts['Weak'] + (c.vwCountsAsWeak ? d.counts['Very Weak'] : 0), o = d.counts['Outstanding'];
    var maxW = Math.floor(c.capWeak * n + 1e-9), maxO = Math.floor(c.capOut * n + 1e-9);
    var extra = (c.extraRules || []).map(function (r) {
      var cnt = 0; (r.cats || []).forEach(function (k) { cnt += d.counts[k] || 0; });
      var lim = r.op === 'min' ? Math.ceil(r.share * n - 1e-9) : Math.floor(r.share * n + 1e-9);
      return { label: r.label || ruleLabel(r), count: cnt, limit: lim, op: r.op, ok: r.op === 'min' ? cnt >= lim : cnt <= lim };
    });
    var extraOk = extra.every(function (x) { return x.ok; });
    return { d: d, n: n, weak: w, out: o, maxW: maxW, maxO: maxO, weakOk: w <= maxW, outOk: o <= maxO, extra: extra, extraOk: extraOk, allOk: w <= maxW && o <= maxO && extraOk };
  }

  /* ---------- corrections ---------- */
  function findFixes(parsed, c) {
    var fixes = [], okRows = 0, totalRows = 0, built = build(parsed, c);
    parsed.blocks.forEach(function (b) {
      var tag = parsed.blocks.length > 1 ? b.title + ': ' : '';
      var mis = b.rows.filter(function (x) { return x.label != null && x.phase != null && x.label !== x.phase && b.totals.length; });
      if (mis.length) fixes.push({ kind: 'label', what: tag + 'Phase labels out of place', detail: mis.map(function (x) { return x.grade + ' is labelled Phase ' + x.label + ' but counted in the Phase ' + x.phase + ' total'; }).join('. ') + '.', effect: 'No judgement changes', rows: mis.map(function (x) { return x.r; }) });
      b.totals.forEach(function (t) {
        var mem = b.rows.filter(function (x) { return t.members.indexOf(x.r) >= 0; });
        if (!mem.length) return;
        var sE = 0, sG = 0, sH = 0, sA = 0, k = mem.length;
        mem.forEach(function (x) { var A1 = assessedOf(x, c), f1 = function (v) { return v > 1 ? v / 100 : v; }; sE += x.E || 0; sG += x.G != null ? x.G : Math.round(f1(x.I || 0) * A1); sH += x.H != null ? x.H : Math.round(f1(x.J || 0) * A1); sA += A1; });
        var name = tag + phaseLabel(t.phase) + ' total', parts = [], cells = [];
        if (t.E != null && Math.abs(t.E - sE) > 1e-6) { parts.push('student count ' + fmt(t.E) + ' should be ' + sE); cells.push(['E', t.E, sE]); }
        if (t.G != null && Math.abs(t.G - sG) > 1e-6) { parts.push('expected or better ' + fmt(t.G) + (Math.abs(t.G - sG / k) < 0.01 ? ' (an average)' : '') + ' should be ' + sG); cells.push(['G', t.G, sG]); }
        if (t.H != null && Math.abs(t.H - sH) > 1e-6) { parts.push('better than expected ' + fmt(t.H) + (Math.abs(t.H - sH / k) < 0.01 ? ' (an average)' : '') + ' should be ' + sH); cells.push(['H', t.H, sH]); }
        var f1 = function (v) { return v == null ? null : (v > 1 ? v / 100 : v); };
        var avgOf = function (k) { var v = mem.map(function (x) { return f1(x[k]); }).filter(function (v) { return v != null; }); return v.length ? v.reduce(function (a, b) { return a + b; }, 0) / v.length : null; };
        var tI = f1(t.I), tJ = f1(t.J), mI = avgOf('I'), mJ = avgOf('J');
        if (tI != null && sA && Math.abs(tI - sG / sA) > 5e-5) { parts.push('% expected or better ' + pct(tI, 2) + (mI != null && Math.abs(tI - mI) < 5e-4 ? ' (a simple average of the grades)' : '') + ' should be ' + pct(sG / sA, 2)); cells.push(['I', t.I, sG / sA]); }
        if (tJ != null && sA && Math.abs(tJ - sH / sA) > 5e-5) { parts.push('% better than expected ' + pct(tJ, 2) + (mJ != null && Math.abs(tJ - mJ) < 5e-4 ? ' (a simple average of the grades)' : '') + ' should be ' + pct(sH / sA, 2)); cells.push(['J', t.J, sH / sA]); }
        var ct = built.items.filter(function (it) { return it.kind === 'total' && it.blockId === b.id && it.r === t.r; })[0];
        var eff = !ct ? '' : !t.K ? 'Judgement ' + ct.judg : t.K === ct.judg ? 'None (stays ' + ct.judg + ')' : 'Changes from ' + t.K + ' to ' + ct.judg;
        if (parts.length) fixes.push({ kind: 'total', what: name + ' added up wrong', detail: parts.join('; ') + '.', effect: eff, r: t.r, cells: cells });
      });
      b.rows.forEach(function (x) {
        totalRows++;
        var A = assessedOf(x, c), probs = [], cells = [];
        if (x.G == null || x.H == null) { okRows++; return; }
        if (x.I != null && A && Math.abs(x.I - x.G / A) > 5e-5) { probs.push('% expected or better ' + pct(x.I, 2) + ' should be ' + pct(x.G / A, 2)); cells.push(['I', x.I, x.G / A]); }
        if (x.J != null && A && Math.abs(x.J - x.H / A) > 5e-5) { probs.push('% better than expected ' + pct(x.J, 2) + ' should be ' + pct(x.H / A, 2)); cells.push(['J', x.J, x.H / A]); }
        var nj = judge(A, x.G, x.H, c);
        if (x.K && x.K !== nj) { probs.push('judgement ' + x.K + ' should be ' + nj); cells.push(['K', x.K, nj]); }
        if (x.H > x.G) probs.push('more students did better than expected than made expected progress');
        if (x.G > A) probs.push('more students made progress than were assessed');
        if (probs.length) fixes.push({ kind: 'grade', what: tag + x.grade + ' has a mistake', detail: probs.join('; ') + '.', effect: '', r: x.r, cells: cells });
        else okRows++;
      });
    });
    parsed.blocks.forEach(function (b) {
      if ((b.col.phase == null || b.rows.every(function (x) { return x.label == null; })) && !b.totals.length && b.rows.length > 1) fixes.push({ kind: 'phase', what: (parsed.blocks.length > 1 ? b.title + ': ' : '') + 'No phases in the sheet', detail: 'The table has no phase column and no total rows, so all ' + b.rows.length + ' grades are treated as one phase for now. Set the phase of each grade if the school splits them.', effect: 'Phase results may change' });
    });
    var derivedRows = 0; parsed.blocks.forEach(function (b) { b.rows.forEach(function (x) { if (x.G == null || x.H == null) derivedRows++; }); });
    if (derivedRows) fixes.push({ kind: 'derived', what: 'Counts worked out from percentages', detail: derivedRows + ' row' + (derivedRows === 1 ? '' : 's') + ' gave percentages but not student counts, so the counts were worked out as percentage x assessed students, rounded to whole students.', effect: 'None' });
    if (parsed.hasFormulas === false) fixes.push({ kind: 'typed', what: 'Every result was typed in', detail: 'The sheet had no formulas, so nothing updated when a number changed. The new workbook uses formulas.', effect: '' });
    return { fixes: fixes, okRows: okRows, totalRows: totalRows };
  }
  function fmt(v) { return Math.abs(v - Math.round(v)) < 1e-9 ? String(Math.round(v)) : (Math.round(v * 100) / 100).toString(); }
  function pct(v, d) { return v == null ? '–' : (v * 100).toFixed(d == null ? 1 : d) + '%'; }

  /* ---------- simulation ---------- */
  function subsets(arr, max) {
    var out = [], n = Math.min(arr.length, max || 12);
    for (var m = 0; m < (1 << n); m++) { var s = []; for (var i = 0; i < n; i++) if (m & (1 << i)) s.push(arr[i]); out.push(s); }
    return out;
  }
  // Options for one unit: which of its bad grades to move, and the resulting judgements
  function unitOptions(parsed, c, u, base, dir) {
    var locked = c.simLocked || [];
    var bad = u.grades.filter(function (g) { return locked.indexOf(g.id) < 0 && (dir === 'up' ? isWeakish(g.judg, c) : g.judg === 'Outstanding'); });
    function capOk(g, n) { if (c.simMaxStudents != null && n > c.simMaxStudents) return false; if (c.simMaxPts != null && g.A && n / g.A * 100 > c.simMaxPts + 1e-9) return false; return true; }
    var opts = [];
    subsets(bad, 12).forEach(function (S) {
      var adds = {}, students = 0, ok = true;
      S.forEach(function (g) {
        if (dir === 'up') { var nd = needAt(g.A, c.acc) - g.G; if (g.G + nd > g.A || !capOk(g, nd)) ok = false; adds[g.id] = { G: g.G + nd }; students += nd; }
        else { var dn = g.H - (needAt(g.A, c.out) - 1); if (!capOk(g, dn)) ok = false; adds[g.id] = { H: g.H - dn }; students += dn; }
      });
      if (!ok) return;
      opts.push({ adds: adds, students: students, extras: 0, moved: S.map(function (g) { return g.id; }) });
      // Totals judged from counts may need extra students to cross the line on their own
      if (dir === 'up' && c.totalMethod === 'counts' && u.total && isWeakish(u.total.judg, c)) {
        var sumG = u.total.G + students, need = needAt(u.total.A, c.acc) - sumG;
        if (need > 0) {
          var a2 = JSON.parse(JSON.stringify(adds)), left = need;
          var free = u.grades.filter(function (g) { return locked.indexOf(g.id) < 0; });
          var order = S.concat(free.filter(function (g) { return S.indexOf(g) < 0 && !isWeakish(g.judg, c); }), free.filter(function (g) { return S.indexOf(g) < 0 && isWeakish(g.judg, c); }));
          order.forEach(function (g) {
            if (left <= 0) return;
            var cur = a2[g.id] ? a2[g.id].G : g.G;
            var cap = S.indexOf(g) >= 0 || !isWeakish(g.judg, c) ? g.A - cur : Math.max(0, needAt(g.A, c.acc) - 1 - cur);
            if (c.simMaxStudents != null) cap = Math.min(cap, c.simMaxStudents - (cur - g.G));
            if (c.simMaxPts != null && g.A) cap = Math.min(cap, Math.floor(c.simMaxPts / 100 * g.A + 1e-9) - (cur - g.G));
            var take = Math.min(cap, left); if (take > 0) { a2[g.id] = { G: cur + take }; left -= take; }
          });
          if (left === 0) opts.push({ adds: a2, students: students + need, extras: need, moved: S.map(function (g) { return g.id; }) });
        }
      }
    });
    return opts;
  }
  function evaluate(parsed, c, overrides) { return build(parsed, c, overrides); }
  function key(id, parsed) { return id; }
  function simulate(parsed, c, maxOptions) {
    maxOptions = maxOptions || 6;
    var base = build(parsed, c);
    var chk0 = check(counted(base.items, c.basis), c);
    var pinned = {}, ids = {}; base.items.forEach(function (it) { ids[it.id] = it; });
    Object.keys(c.simPinned || {}).forEach(function (k) { if (ids[k] && ids[k].kind === 'grade') pinned[k] = Object.assign({}, c.simPinned[k]); });
    var hasPinned = Object.keys(pinned).length > 0;
    if (chk0.weakOk && chk0.outOk && !hasPinned) return { needed: false, base: base, check: chk0 };
    var overrides = JSON.parse(JSON.stringify(pinned));
    var chkP = check(counted(build(parsed, c, overrides).items, c.basis), c);
    if (hasPinned && chkP.weakOk && chkP.outOk) { var only = [{ overrides: overrides, extras: 0 }]; return describe(only, true); }
    // Outstanding first (rare), then Weak
    var stages = [];
    if (!chk0.outOk) stages.push('down');
    stages.push('up');
    var finalOpts = null, cur = hasPinned ? build(parsed, c, overrides) : base;
    if (!chkP.outOk) { if (stages[0] !== 'down') stages.unshift('down'); } else if (stages[0] === 'down') stages.shift();
    for (var si = 0; si < stages.length; si++) {
      var dir = stages[si];
      var chk = check(counted(cur.items, c.basis), c);
      var needRed = dir === 'up' ? chk.weak - chk.maxW : chk.out - chk.maxO;
      if (needRed <= 0) { if (dir === 'up') finalOpts = [{ overrides: overrides }]; continue; }
      // evaluate every unit option once against the current state
      var perUnit = cur.units.map(function (u) {
        return unitOptions(parsed, c, u, cur, dir).map(function (o) {
          var ov = Object.assign({}, overrides); Object.keys(o.adds).forEach(function (k) { ov[k] = Object.assign({}, ov[k] || {}, o.adds[k]); });
          var m = build(parsed, c, ov), uu = m.units.filter(function (x) { return x.key === u.key; })[0];
          var before = u.grades.concat(u.total ? [u.total] : []), after = uu.grades.concat(uu.total ? [uu.total] : []);
          var changes = 0, red = 0, newOut = 0;
          after.forEach(function (it, i) {
            if (it.judg !== before[i].judg) changes++;
            var counts = c.basis === 'both' || (c.basis === 'grades' ? it.kind === 'grade' : it.kind === 'total');
            if (!counts) return;
            if (dir === 'up') red += (isWeakish(before[i].judg, c) ? 1 : 0) - (isWeakish(it.judg, c) ? 1 : 0);
            else red += (before[i].judg === 'Outstanding' ? 1 : 0) - (it.judg === 'Outstanding' ? 1 : 0);
            if (dir === 'up') newOut += (it.judg === 'Outstanding' ? 1 : 0) - (before[i].judg === 'Outstanding' ? 1 : 0);
          });
          return { adds: o.adds, students: o.students, extras: o.extras, changes: changes, red: red, newOut: newOut, unit: u.key };
        }).filter(function (o) { return o.red > 0 || Object.keys(o.adds).length === 0; });
      });
      // K-best dynamic programme over units: state = weak (or outstanding) reduction achieved, capped at need
      var K = Math.max(maxOptions * 4, 24);
      var dp = {}; dp[0] = [{ changes: 0, students: 0, extras: 0, picks: [] }];
      perUnit.forEach(function (opts) {
        var nd = {};
        Object.keys(dp).forEach(function (k) {
          dp[k].forEach(function (s) {
            opts.forEach(function (o) {
              var nk = Math.min(needRed, +k + Math.max(0, o.red));
              var cand = { changes: s.changes + o.changes, students: s.students + o.students, extras: s.extras + o.extras, picks: Object.keys(o.adds).length ? s.picks.concat([o]) : s.picks };
              (nd[nk] = nd[nk] || []).push(cand);
            });
          });
        });
        Object.keys(nd).forEach(function (k) {
          nd[k].sort(function (a, b) { return a.changes - b.changes || a.students - b.students || a.extras - b.extras; });
          var seen = {}, keep = [];
          nd[k].forEach(function (s) { var sig = s.picks.map(function (p) { return Object.keys(p.adds).sort().join(',') + '#' + p.students; }).sort().join('|'); if (!seen[sig] && keep.length < K) { seen[sig] = 1; keep.push(s); } });
          nd[k] = keep;
        });
        dp = nd;
      });
      var sols = (dp[needRed] || []).map(function (s) {
        var ov = Object.assign({}, overrides);
        s.picks.forEach(function (p) { Object.keys(p.adds).forEach(function (k) { ov[k] = Object.assign({}, ov[k] || {}, p.adds[k]); }); });
        return { overrides: ov, changes: s.changes, students: s.students, extras: s.extras };
      });
      // verify against the full rules and keep the valid ones
      sols = sols.filter(function (s) { var m = build(parsed, c, s.overrides), k2 = check(counted(m.items, c.basis), c); return dir === 'up' ? (k2.weakOk && k2.outOk) : k2.outOk; });
      if (!sols.length) return { needed: true, failed: true, base: base, check: chk0, reason: dir === 'up' ? 'weak' : 'out', limited: !!((c.simLocked || []).length || c.simMaxStudents != null || c.simMaxPts != null) };
      if (dir === 'up') { var withExtra = sols.filter(function (s) { var m = build(parsed, c, s.overrides); return check(counted(m.items, c.basis), c).extraOk; }); if (withExtra.length) sols = withExtra; }
      if (dir === 'down') { overrides = sols[0].overrides; cur = build(parsed, c, overrides); }
      else finalOpts = sols;
    }
    return describe(finalOpts || [], hasPinned);
    function describe(finalOpts, hasPinned) {
    var options = finalOpts.slice(0, maxOptions).map(function (s) {
      var m = build(parsed, c, s.overrides);
      var cells = [], changed = [], students = 0, maxPts = 0;
      m.items.forEach(function (it, i) {
        var b = base.items[i];
        if (it.judg !== b.judg) changed.push({ from: b, to: it });
        if (it.kind === 'grade' && (it.G !== b.G || it.H !== b.H)) {
          var dG = it.G - b.G, dH = it.H - b.H; students += Math.abs(dG) + Math.abs(dH);
          var pp = Math.max(Math.abs((it.I - b.I) * 100), Math.abs((it.J - b.J) * 100)); maxPts = Math.max(maxPts, pp);
          cells.push({ from: b, to: it, dG: dG, dH: dH, pp: pp });
        }
      });
      var assessedAll = base.items.filter(function (x) { return x.kind === 'grade'; }).reduce(function (a, x) { return a + (x.A || 0); }, 0);
      return { overrides: s.overrides, model: m, changed: changed, cells: cells, students: students, extras: s.extras, maxPts: maxPts,
        share: assessedAll ? students / assessedAll : 0, substantial: maxPts > c.substantialPts + 1e-9, check: check(counted(m.items, c.basis), c) };
    });
    return { needed: true, base: base, check: chk0, options: options, total: finalOpts.length, pinned: hasPinned, extraFail: options.length && !options[0].check.extraOk };
    }
  }

  return { CATS: CATS, ROLES: ROLES, ROLE_LABEL: ROLE_LABEL, defaults: defaults, needAt: needAt, judge: judge, roundBy: roundBy, catFromPoints: catFromPoints,
    gridFromAoa: gridFromAoa, parse: parse, headerOptions: headerOptions, colLetter: colLetter, assessedOf: assessedOf,
    build: build, counted: counted, ruleLabel: ruleLabel, phaseLabel: phaseLabel, dist: dist, check: check, findFixes: findFixes, simulate: simulate, isWeakish: isWeakish, pct: pct, fmt: fmt };
})();
if (typeof module !== 'undefined') module.exports = Engine;
