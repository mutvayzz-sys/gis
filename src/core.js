/* Core logic and HTML renderers. Pure functions returning strings, so the same code
   prerenders the page at build time and re-renders it live in the browser. */
var Core = (function () {
  var CATS = ['Weak', 'Acceptable', 'Good', 'Very Good', 'Outstanding'];
  var CAT_KEY = { 'Weak': 'weak', 'Acceptable': 'acc', 'Good': 'good', 'Very Good': 'vg', 'Outstanding': 'out' };
  var DEFAULT_RULES = { acc: 0.75, good: 0.50, vg: 0.61, out: 0.75, capWeak: 0.20, capOut: 0.20, mode: 'all' };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function pct(x, d) { if (x == null || isNaN(x)) return '–'; return (x * 100).toFixed(d == null ? 1 : d) + '%'; }
  function num(x) {
    if (x == null || x === '') return '–';
    if (Math.abs(x - Math.round(x)) < 1e-9) return String(Math.round(x));
    return (Math.round(x * 100) / 100).toString();
  }
  function needAt(E, t) { return Math.max(0, Math.ceil(t * E - 1e-9)); }
  function judge(E, G, H, r) {
    if (!(E > 0)) return 'No data';
    if (G < needAt(E, r.acc)) return 'Weak';
    if (H >= needAt(E, r.out)) return 'Outstanding';
    if (H >= needAt(E, r.vg)) return 'Very Good';
    if (H >= needAt(E, r.good)) return 'Good';
    return 'Acceptable';
  }
  function gradeName(g) {
    var m = /^\s*(?:g|grade|y|year)\s*([0-9]+)\s*$/i.exec(String(g));
    return m ? 'Grade ' + m[1] : String(g);
  }

  /* ---------- parsing ---------- */
  function toNum(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    var s = String(v).replace(/[,%\s]/g, '');
    var n = parseFloat(s);
    return isNaN(n) ? null : n;
  }
  function parseGrid(aoa, meta) {
    meta = meta || {};
    var hdr = -1, col = {};
    for (var r = 0; r < Math.min(aoa.length, 30) && hdr < 0; r++) {
      var row = aoa[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (row[c] != null && /grade|year/i.test(String(row[c])) && !/%/.test(String(row[c]))) { hdr = r; break; }
      }
    }
    if (hdr >= 0) {
      (aoa[hdr] || []).forEach(function (h, c) {
        if (h == null) return;
        var s = String(h).toLowerCase();
        if (/phase/.test(s)) col.phase = c;
        else if (/grade|year/.test(s)) col.grade = c;
        else if (/no\.? of students|students in/.test(s)) col.E = c;
        else if (/assessed/.test(s)) col.F = c;
        else if (/%/.test(s) && /expected or better/.test(s)) col.I = c;
        else if (/%/.test(s) && /better than expected/.test(s)) col.J = c;
        else if (/expected or better/.test(s)) col.G = c;
        else if (/better than expected/.test(s)) col.H = c;
        else if (/judg/.test(s)) col.K = c;
      });
    }
    if (hdr < 0 || col.grade == null || col.E == null || col.G == null || col.H == null) {
      // Plain paste: Grade, Phase, Students, Expected or better, Better than expected
      hdr = -1; col = { grade: 0, phase: 1, E: 2, G: 3, H: 4 };
    }
    var rows = [], totals = [], block = [], blockExplicit = null, carry = null, raw = [];
    for (var i = hdr + 1; i < aoa.length; i++) {
      var a = aoa[i] || [];
      var g = a[col.grade];
      if (g == null || String(g).trim() === '') continue;
      var lab = col.phase != null ? toNum(a[col.phase]) : null;
      if (lab != null) { carry = lab; if (blockExplicit == null) blockExplicit = lab; }
      var rec = {
        grade: String(g).trim(), label: carry, E: toNum(a[col.E]), G: toNum(a[col.G]), H: toNum(a[col.H]),
        F: col.F != null ? toNum(a[col.F]) : null,
        I: col.I != null ? toNum(a[col.I]) : null, J: col.J != null ? toNum(a[col.J]) : null,
        K: col.K != null && a[col.K] != null ? String(a[col.K]).trim() : null
      };
      raw.push(rec);
      if (/total/i.test(rec.grade)) {
        var ph = blockExplicit != null ? blockExplicit : (block[0] ? block[0].label : totals.length + 1);
        block.forEach(function (b) { b.phase = ph; });
        rec.phase = ph; rec.members = block.map(function (b) { return b.grade; });
        totals.push(rec); block = []; blockExplicit = null;
      } else {
        if (rec.E == null || rec.G == null || rec.H == null) continue;
        rows.push(rec); block.push(rec);
      }
    }
    block.forEach(function (b) { b.phase = b.label != null ? b.label : (totals.length ? totals[totals.length - 1].phase + 1 : 1); });
    var subject = null;
    for (var tr = 0; tr < Math.max(0, hdr); tr++) {
      (aoa[tr] || []).forEach(function (v) { if (!subject && v != null && typeof v === 'string' && v.trim() && v.length < 60) subject = v.trim(); });
    }
    return {
      subject: subject, source: meta.source || 'Pasted data', hasFormulas: meta.hasFormulas, raw: raw, totals: totals,
      rows: rows.map(function (r) { return { grade: r.grade, phase: r.phase, E: r.E, G: r.G, H: r.H, label: r.label, F: r.F, I: r.I, J: r.J, K: r.K }; })
    };
  }
  function parseText(text) {
    var lines = String(text).replace(/\r/g, '').split('\n').filter(function (l) { return l.trim() !== ''; });
    var sep = lines.some(function (l) { return l.indexOf('\t') >= 0; }) ? '\t' : ',';
    return lines.map(function (l) { return l.split(sep).map(function (v) { v = v.trim(); return v === '' ? null : v; }); });
  }

  /* ---------- model ---------- */
  function build(rows, r) {
    var phases = [], byPhase = {};
    rows.forEach(function (row, idx) {
      var key = String(row.phase);
      if (!byPhase[key]) { byPhase[key] = { phase: row.phase, grades: [] }; phases.push(byPhase[key]); }
      byPhase[key].grades.push({ id: 'g' + idx, idx: idx, grade: row.grade, phase: row.phase, E: +row.E, G: +row.G, H: +row.H });
    });
    var items = [];
    phases.forEach(function (p) {
      p.grades.forEach(function (g) { finish(g, r); g.kind = 'grade'; g.label = gradeName(g.grade); items.push(g); });
      if (p.grades.length >= 2) {
        var t = { id: 't' + p.phase, kind: 'total', phase: p.phase, label: 'Phase ' + p.phase + ' total', E: 0, G: 0, H: 0 };
        p.grades.forEach(function (g) { t.E += g.E; t.G += g.G; t.H += g.H; });
        finish(t, r); p.total = t; items.push(t);
      }
    });
    return { phases: phases, items: items };
  }
  function finish(o, r) { o.I = o.E > 0 ? o.G / o.E : null; o.J = o.E > 0 ? o.H / o.E : null; o.judg = judge(o.E, o.G, o.H, r); }
  function counted(items, mode) {
    return items.filter(function (it) { return mode === 'all' || (mode === 'grades' ? it.kind === 'grade' : it.kind === 'total'); });
  }
  function dist(list) {
    var c = {}; CATS.forEach(function (k) { c[k] = 0; });
    list.forEach(function (it) { if (c[it.judg] != null) c[it.judg]++; });
    return { counts: c, n: list.length };
  }
  function meets(d, r) {
    var n = d.n || 1;
    return { weak: d.counts['Weak'] / n <= r.capWeak + 1e-9, out: d.counts['Outstanding'] / n <= r.capOut + 1e-9 };
  }

  /* ---------- simulation search ---------- */
  // dir 'up': raise G (expected or better) to lift Weak; dir 'down': lower H to bring Outstanding down to Very Good.
  function search(rows, r, dir) {
    var m = build(rows, r);
    var badJ = dir === 'up' ? 'Weak' : 'Outstanding';
    var cands = [];
    m.phases.forEach(function (p) {
      p.grades.forEach(function (g) { if (g.judg === badJ) cands.push({ type: 'g', g: g }); });
      if (p.total && p.total.judg === badJ) cands.push({ type: 't', p: p });
    });
    var bits = cands.length;
    if (bits > 22) return { tooMany: true, options: [] };
    var base = counted(m.items, r.mode);
    var results = [];
    for (var mask = 0; mask < (1 << bits); mask++) {
      var flip = {}, flipT = {}, nflip = 0;
      for (var b = 0; b < bits; b++) if (mask & (1 << b)) { nflip++; if (cands[b].type === 'g') flip[cands[b].g.id] = 1; else flipT[cands[b].p.phase] = 1; }
      var adds = {}, ok = true, students = 0, extras = 0;
      for (var pi = 0; pi < m.phases.length && ok; pi++) {
        var p = m.phases[pi], lo = {}, hi = {}, smin = 0, smax = 0;
        p.grades.forEach(function (g) {
          var l, h;
          if (dir === 'up') {
            var nd = needAt(g.E, r.acc);
            if (g.judg === 'Weak') { if (flip[g.id]) { l = nd - g.G; h = g.E - g.G; } else { l = 0; h = nd - 1 - g.G; } }
            else { l = 0; h = g.E - g.G; }
          } else {
            var no = needAt(g.E, r.out), nv = needAt(g.E, r.vg);
            if (g.judg === 'Outstanding') { if (flip[g.id]) { l = g.H - (no - 1); h = Math.max(l, g.H - nv); } else { l = 0; h = g.H - no; } }
            else { l = 0; h = 0; }
          }
          lo[g.id] = l; hi[g.id] = h; smin += l; smax += h;
        });
        var s = smin;
        if (p.total && p.total.judg === badJ) {
          var req = dir === 'up' ? needAt(p.total.E, r.acc) - p.total.G : p.total.H - (needAt(p.total.E, r.out) - 1);
          if (flipT[p.phase]) { s = Math.max(smin, req); if (s > smax) ok = false; }
          else if (s > req - 1) ok = false;
        }
        if (!ok) break;
        // allocate: minimum first, extra to grades already changing, then others
        p.grades.forEach(function (g) { adds[g.id] = lo[g.id]; });
        var extra = s - smin; extras += extra;
        var groups = [p.grades.filter(function (g) { return flip[g.id]; }), p.grades.filter(function (g) { return !flip[g.id] && g.judg === badJ && adds[g.id] > 0; }), p.grades.filter(function (g) { return !flip[g.id] && g.judg === badJ && !(adds[g.id] > 0); }), p.grades.filter(function (g) { return g.judg !== badJ; })];
        for (var gi = 0; gi < groups.length && extra > 0; gi++) {
          var order = groups[gi], guard = 0;
          while (extra > 0 && guard < 100000) {
            var moved = false;
            for (var oi = 0; oi < order.length && extra > 0; oi++) {
              var gg = order[oi];
              if (adds[gg.id] < hi[gg.id]) { adds[gg.id]++; extra--; moved = true; }
            }
            if (!moved) break;
            guard++;
          }
        }
        if (extra > 0) ok = false;
        students += s;
      }
      if (!ok) continue;
      var newRows = applyAdds(rows, m, adds, dir);
      var nm = build(newRows, r);
      var changed = [];
      nm.items.forEach(function (it, k) { if (it.judg !== m.items[k].judg) changed.push({ from: m.items[k], to: it }); });
      var d = dist(counted(nm.items, r.mode));
      var mt = meets(d, r);
      if (!(dir === 'up' ? mt.weak : mt.out)) continue;
      var cells = 0; Object.keys(adds).forEach(function (k) { if (adds[k] > 0) cells++; });
      results.push({ changes: changed.length, students: students, extras: extras, cells: cells, adds: adds, rows: newRows, model: nm, changed: changed });
    }
    if (!results.length) return { options: [], impossible: true };
    var minC = Math.min.apply(null, results.map(function (x) { return x.changes; }));
    var best = results.filter(function (x) { return x.changes === minC; });
    best.sort(function (a, b) { return a.students - b.students || a.extras - b.extras || a.cells - b.cells; });
    var seen = {}, uniq = [];
    best.forEach(function (x) { var sig = JSON.stringify(x.adds); if (!seen[sig]) { seen[sig] = 1; uniq.push(x); } });
    return { options: uniq, minChanges: minC, baseModel: m };
  }
  function applyAdds(rows, m, adds, dir) {
    var out = rows.map(function (x) { return Object.assign({}, x); });
    m.phases.forEach(function (p) {
      p.grades.forEach(function (g) {
        var a = adds[g.id] || 0;
        if (dir === 'up') out[g.idx].G = g.G + a; else out[g.idx].H = g.H - a;
      });
    });
    return out;
  }
  function simulate(rows, r) {
    var m = build(rows, r);
    var d0 = dist(counted(m.items, r.mode));
    var mt = meets(d0, r);
    if (mt.weak && mt.out) return { needed: false, model: m };
    var workRows = rows, outPlan = null;
    if (!mt.out) {
      var so = search(rows, r, 'down');
      if (!so.options.length) return { needed: true, failed: true, model: m };
      outPlan = so.options[0]; workRows = outPlan.rows;
    }
    var opts;
    var mt2 = meets(dist(counted(build(workRows, r).items, r.mode)), r);
    if (!mt2.weak) {
      var sw = search(workRows, r, 'up');
      if (sw.tooMany) return { needed: true, failed: true, tooMany: true, model: m };
      if (!sw.options.length) return { needed: true, failed: true, model: m };
      opts = sw.options;
    } else {
      opts = [{ changes: 0, students: 0, extras: 0, cells: 0, adds: {}, rows: workRows, model: build(workRows, r), changed: [] }];
    }
    var options = opts.map(function (o) {
      var fm = build(o.rows, r);
      var changed = [];
      fm.items.forEach(function (it, k) { if (it.judg !== m.items[k].judg) changed.push({ from: m.items[k], to: it }); });
      var cellChanges = [];
      fm.items.forEach(function (it, k) {
        var b = m.items[k];
        if (it.G !== b.G || it.H !== b.H) cellChanges.push({ from: b, to: it });
      });
      var students = 0, lowered = 0;
      cellChanges.forEach(function (c) { if (c.from.kind === 'grade') { students += c.to.G - c.from.G; lowered += c.from.H - c.to.H; } });
      return { model: fm, rows: o.rows, changed: changed, cellChanges: cellChanges, students: students, lowered: lowered, extras: o.extras + (outPlan ? outPlan.extras : 0) };
    });
    var minStudents = Math.min.apply(null, options.map(function (o) { return o.students; }));
    return { needed: true, model: m, options: options, total: options.length, minStudents: minStudents };
  }

  /* ---------- renderers ---------- */
  function dot(cat) { return '<span class="dot c-' + (CAT_KEY[cat] || 'na') + '" aria-hidden="true"></span>'; }
  function judgTag(cat) { return '<span class="judg">' + dot(cat) + esc(cat) + '</span>'; }

  function shortLabel(it) { return it.kind === 'total' ? 'P' + it.phase + ' total' : String(it.grade).replace(/^grade\s*/i, 'G'); }
  function plural(n, one, many) { return n + ' ' + (n === 1 ? one : (many || one + 's')); }

  // One box per judgement, sorted Weak first, with the allowed Weak spots marked.
  function renderDist(list, r, opts) {
    opts = opts || {};
    var d = dist(list), n = d.n || 1, mt = meets(d, r);
    var maxW = Math.floor(r.capWeak * d.n + 1e-9), maxO = Math.floor(r.capOut * d.n + 1e-9);
    var w = d.counts['Weak'], o = d.counts['Outstanding'];
    var order = {}; CATS.forEach(function (k, i) { order[k] = i; });
    var sorted = list.slice().sort(function (a, b) { return (order[a.judg] - order[b.judg]) || 0; });
    function tile(it, k) {
      var over = it.judg === 'Weak' && k >= maxW;
      return '<li class="tile' + (over ? ' over' : '') + '"><span class="sq c-' + (CAT_KEY[it.judg] || 'na') + '" aria-hidden="true"></span><span class="tl">' + esc(it.label ? shortLabel(it) : '') + '</span></li>';
    }
    var allowed = sorted.slice(0, maxW), rest = sorted.slice(maxW);
    var tilesHtml = '<div class="tiles" role="img" aria-label="' + esc(CATS.map(function (k) { return d.counts[k] + ' ' + k; }).join(', ')) + '">' +
      (maxW > 0 ? '<div class="allowed"><span class="allowed-cap">' + plural(maxW, 'spot') + ' allowed to be Weak</span><ul>' + allowed.map(tile).join('') + '</ul></div>' : '<div class="allowed none"><span class="allowed-cap">No spots allowed to be Weak</span></div>') +
      '<ul class="rest">' + rest.map(function (it, k) { return tile(it, k + maxW); }).join('') + '</ul></div>';
    var chips = '<ul class="chips">' + CATS.map(function (k) {
      return '<li>' + dot(k) + '<strong>' + d.counts[k] + '</strong> ' + esc(k) + '</li>';
    }).join('') + '</ul>';
    var weakLine = mt.weak
      ? '<li class="check is-ok"><span class="dot ok" aria-hidden="true"></span><span><strong>Weak is fine.</strong> ' + w + ' of ' + d.n + ' are Weak, and up to ' + maxW + ' are allowed.</span></li>'
      : '<li class="check is-bad"><span class="dot bad" aria-hidden="true"></span><span><strong>Too many Weak.</strong> ' + w + ' of ' + d.n + ' are Weak, but only ' + maxW + ' are allowed. That is ' + (w - maxW) + ' too many. The red boxes outside the dashed area are the extra ones.</span></li>';
    var outLine = mt.out
      ? '<li class="check is-ok"><span class="dot ok" aria-hidden="true"></span><span><strong>Outstanding is fine.</strong> ' + o + ' of ' + d.n + ' are Outstanding, and up to ' + maxO + ' are allowed.</span></li>'
      : '<li class="check is-bad"><span class="dot bad" aria-hidden="true"></span><span><strong>Too many Outstanding.</strong> ' + o + ' of ' + d.n + ', but only ' + maxO + ' are allowed.</span></li>';
    return '<div class="dist">' + (opts.intro !== false ? '<p class="dist-intro">Each box is one judgement. Red means Weak.</p>' : '') + tilesHtml + chips + '<ul class="checks">' + weakLine + outLine + '</ul></div>';
  }

  function pctNice(x) {
    if (x == null || isNaN(x)) return '–';
    var v = x * 100;
    return (Math.abs(v - Math.round(v)) < 1e-6 ? Math.round(v) : v.toFixed(1)) + '%';
  }
  function bar(v, t) {
    if (v == null) return '';
    var w = Math.max(0, Math.min(1, v)) * 100;
    return '<span class="meter' + (v + 1e-9 >= t ? ' pass' : '') + '" aria-hidden="true"><span style="width:' + w.toFixed(2) + '%"></span><i style="left:' + (t * 100) + '%"></i></span>';
  }
  function needText(it, r) {
    if (it.judg === 'Weak') {
      var k = needAt(it.E, r.acc) - it.G;
      return '<span class="need bad">' + plural(k, 'more student') + (it.kind === 'total' ? ' across its grades' : '') + '</span>';
    }
    if (it.judg === 'No data') return '';
    return '<span class="need ok">Reached</span>';
  }
  function renderTable(model, r, opts) {
    opts = opts || {};
    var base = opts.base, mode = r.mode;
    var target = pctNice(r.acc);
    var head = '<thead><tr><th scope="col">Grade</th><th scope="col">Made expected progress<span class="th-sub">target ' + target + ', the black line</span></th><th scope="col">Did better than expected</th><th scope="col">Result</th><th scope="col">To reach ' + target + '</th></tr></thead>';
    var body = '';
    var lookup = {};
    if (base) base.items.forEach(function (it) { lookup[it.id] = it; });
    model.phases.forEach(function (p) {
      var cnt = p.grades.length;
      body += '<tr class="ph-row"><th colspan="5" scope="colgroup">Phase ' + esc(p.phase) + '<span class="ph-sub">' + esc(gradeName(p.grades[0].grade)) + (cnt > 1 ? ' to ' + esc(gradeName(p.grades[cnt - 1].grade)) : '') + '</span></th></tr>';
      var list = p.grades.slice(); if (p.total) list.push(p.total);
      list.forEach(function (it) {
        var b = lookup[it.id];
        var isCounted = mode === 'all' || (mode === 'grades' ? it.kind === 'grade' : it.kind === 'total');
        function frac(field, withBar) {
          var changed = b && b[field] !== it[field];
          var dlt = changed ? it[field] - b[field] : 0;
          return '<td class="pc' + (changed ? ' chg' : '') + '"><div class="pc-line"><span class="pc-n"><strong>' + num(it[field]) + '</strong> of ' + num(it.E) + '</span><span class="pc-p">' + pctNice(field === 'G' ? it.I : it.J) + '</span></div>' +
            (withBar ? bar(it.I, r.acc) : '') +
            (changed ? '<span class="pc-was">' + (dlt > 0 ? '+' : '') + num(dlt) + ' students, was ' + num(b[field]) + '</span>' : '') + '</td>';
        }
        var name = it.kind === 'total' ? 'Phase ' + it.phase + ' overall' : gradeName(it.grade);
        var sub = it.kind === 'total' ? '<span class="row-sub">all its grades added up</span>' : '';
        var jc = '<td class="jc">' + judgTag(it.judg) + (b && b.judg !== it.judg ? '<span class="was-j">was ' + esc(b.judg) + '</span>' : '') + (!isCounted ? '<span class="nc">not counted</span>' : '') + '</td>';
        body += '<tr class="' + (it.kind === 'total' ? 'tot' : '') + (b && b.judg !== it.judg ? ' moved' : '') + '"><th scope="row">' + esc(name) + sub + '</th>' +
          frac('G', true) + frac('H', false) + jc + '<td class="nd">' + needText(it, r) + '</td></tr>';
      });
    });
    return '<div class="tbl-scroll"><table class="res">' + (opts.caption ? '<caption>' + esc(opts.caption) + '</caption>' : '') + head + '<tbody>' + body + '</tbody></table></div>';
  }

  function findFixes(parsed, r) {
    var fixes = [], fine = [];
    if (!parsed) return { fixes: fixes, fine: fine };
    var rows = parsed.rows;
    var mis = rows.filter(function (x) { return x.label != null && x.phase != null && x.label !== x.phase; });
    if (mis.length && parsed.totals.length) {
      fixes.push({ what: plural(mis.length, 'grade was', 'grades were') + ' filed under the wrong phase', detail: mis.map(function (x) { return gradeName(x.grade) + ' was labelled Phase ' + x.label + ', but its students are counted in Phase ' + x.phase; }).join('. ') + '. The labels were moved to match.', effect: 'No result changed' });
    }
    parsed.totals.forEach(function (t) {
      var mem = rows.filter(function (x) { return t.members.indexOf(x.grade) >= 0; });
      if (!mem.length) return;
      var sE = 0, sG = 0, sH = 0, aI = 0;
      mem.forEach(function (x) { sE += x.E; sG += x.G; sH += x.H; aI += x.G / x.E; });
      var k = mem.length, nm = 'Phase ' + t.phase + ' overall';
      var nj = judge(sE, sG, sH, r);
      var parts = [];
      if (t.E != null && Math.abs(t.E - sE) > 1e-6) parts.push('It said ' + num(t.E) + ' students in total, but the grades add up to ' + sE + '.');
      var avgd = t.G != null && Math.abs(t.G - sG) > 1e-6;
      if (avgd) parts.push('It said ' + num(t.G) + ' students made expected progress. That is the average of the grades, not the total. Added up, it is ' + sG + ' out of ' + sE + '.');
      if (t.H != null && Math.abs(t.H - sH) > 1e-6) parts.push('Did better than expected: ' + num(t.H) + ' should be ' + sH + '.');
      var wI = sG / sE, wJ = sH / sE;
      var pBad = (t.I != null && Math.abs(t.I - wI) > 5e-5) || (t.J != null && Math.abs(t.J - wJ) > 5e-5);
      if (pBad) parts.push('Its percentages treated a small grade the same as a big one. Fixed: ' + pctNice(wI) + ' made expected progress (was ' + pctNice(t.I) + ').');
      if (parts.length) fixes.push({ what: nm + ' was added up wrong', detail: parts.join(' '), effect: t.K && t.K !== nj ? 'Result changes from ' + t.K + ' to ' + nj : 'Result stays ' + nj });
      else if (t.K && t.K !== nj) fixes.push({ what: nm + ' had the wrong result', detail: 'It said ' + t.K + ', but the numbers give ' + nj + '.', effect: '' });
    });
    var okRows = 0;
    rows.forEach(function (x) {
      var jI = x.G / x.E, jJ = x.H / x.E, nj = judge(x.E, x.G, x.H, r), probs = [];
      if (x.I != null && Math.abs(x.I - jI) > 5e-5) probs.push('its percentage for expected progress should be ' + pctNice(jI) + ', not ' + pctNice(x.I));
      if (x.J != null && Math.abs(x.J - jJ) > 5e-5) probs.push('its percentage for better than expected should be ' + pctNice(jJ) + ', not ' + pctNice(x.J));
      if (x.H > x.G) probs.push('more students did better than expected than made expected progress, which cannot happen');
      if (x.G > x.E) probs.push('more students made progress than are in the grade');
      if (x.K && x.K !== nj) probs.push('it said ' + x.K + ', but the numbers give ' + nj);
      if (probs.length) fixes.push({ what: gradeName(x.grade) + ' has a mistake', detail: probs.join('; ').replace(/^./, function (c) { return c.toUpperCase(); }) + '.', effect: '' });
      else okRows++;
    });
    if (parsed.hasFormulas === false) fixes.push({ what: 'The numbers were typed in by hand', detail: 'The sheet had no formulas, so nothing would update if a number changed. This page does the maths itself.', effect: '' });
    if (okRows) fine.push(okRows === rows.length ? 'Every grade row was already right.' : okRows + ' of ' + rows.length + ' grade rows were already right.');
    return { fixes: fixes, fine: fine };
  }
  function renderFixes(parsed, r) {
    var f = findFixes(parsed, r);
    if (!parsed || !parsed.totals) return '';
    var html = '';
    if (!f.fixes.length) html += '<p class="muted">Nothing to fix. Every number and result in the loaded data is right.</p>';
    else html += '<ol class="fixes">' + f.fixes.map(function (x) { return '<li><strong>' + esc(x.what) + '.</strong> ' + esc(x.detail) + (x.effect ? ' <span class="eff">' + esc(x.effect) + '.</span>' : '') + '</li>'; }).join('') + '</ol>';
    if (f.fine.length) html += '<p class="fine">' + f.fine.map(esc).join(' ') + '</p>';
    return html;
  }

  function originalList(parsed, mode) {
    if (!parsed) return null;
    var list = [];
    parsed.rows.forEach(function (x) { if (x.K && mode !== 'totals') list.push({ judg: x.K }); });
    parsed.totals.forEach(function (t) { if (t.K && mode !== 'grades') list.push({ judg: t.K }); });
    return list.length ? list : null;
  }
  function renderCompare(parsed, model, sim, r) {
    var cols = [];
    var o = originalList(parsed, r.mode);
    if (o) cols.push({ name: 'In the file', d: dist(o) });
    cols.push({ name: 'Fixed (real)', d: dist(counted(model.items, r.mode)) });
    if (sim) cols.push({ name: 'Pretend fix', d: dist(counted(sim.model.items, r.mode)), sim: true });
    var h = '<div class="tbl-scroll"><table class="cmp"><thead><tr><th scope="col">Judgement</th>' + cols.map(function (c) { return '<th scope="col" class="n' + (c.sim ? ' simh' : '') + '" colspan="2">' + esc(c.name) + '</th>'; }).join('') + '</tr></thead><tbody>';
    CATS.forEach(function (k) {
      h += '<tr><th scope="row">' + judgTag(k) + '</th>' + cols.map(function (c) { return '<td class="n">' + c.d.counts[k] + '</td><td class="n muted">' + pctNice(c.d.counts[k] / (c.d.n || 1)) + '</td>'; }).join('') + '</tr>';
    });
    h += '<tr class="tot"><th scope="row">Total</th>' + cols.map(function (c) { return '<td class="n">' + c.d.n + '</td><td class="n muted">100%</td>'; }).join('') + '</tr>';
    h += '<tr><th scope="row">Few enough Weak?</th>' + cols.map(function (c) { var ok = meets(c.d, r).weak; return '<td colspan="2" class="n"><span class="status ' + (ok ? 'ok' : 'bad') + '">' + (ok ? 'Yes' : 'No') + '</span></td>'; }).join('') + '</tr>';
    h += '<tr><th scope="row">Few enough Outstanding?</th>' + cols.map(function (c) { var ok = meets(c.d, r).out; return '<td colspan="2" class="n"><span class="status ' + (ok ? 'ok' : 'bad') + '">' + (ok ? 'Yes' : 'No') + '</span></td>'; }).join('') + '</tr>';
    return h + '</tbody></table></div>';
  }

  function letter(i) { return String.fromCharCode(65 + i); }
  function optionSummary(o) {
    var grades = o.cellChanges.filter(function (c) { return c.from.kind === 'grade'; });
    return grades;
  }
  function renderOption(sim, i, r, parsed) {
    var o = sim.options[i];
    var base = sim.model;
    var grades = optionSummary(o);
    var autoTotals = o.changed.filter(function (c) { return c.from.kind === 'total'; });
    var stillWeak = counted(o.model.items, r.mode).filter(function (x) { return x.judg === 'Weak'; });
    var d = dist(counted(o.model.items, r.mode));
    var lead = '<p class="opt-lead"><strong>' + o.changed.length + ' judgement' + (o.changed.length === 1 ? '' : 's') + ' change.</strong> ' +
      (o.students ? o.students + ' more student' + (o.students === 1 ? '' : 's') + ' would need to make expected progress' : '') +
      (o.lowered ? (o.students ? ', and ' : '') + o.lowered + ' fewer would be counted as better than expected' : '') + '.</p>';
    var steps = '<ul class="steps">' + grades.map(function (c) {
      var f = c.from, t = c.to, up = t.G !== f.G;
      var field = up ? 'G' : 'H';
      var dn = Math.abs(t[field] - f[field]);
      var nd = needAt(f.E, up ? r.acc : r.out);
      var exact = up ? (t.G === nd) : (t.H === nd - 1);
      var txt = up
        ? dn + ' more of its ' + f.E + ' students make expected progress (' + f.G + ' to ' + t.G + '). That moves it from ' + pctNice(f.I) + ' to ' + pctNice(t.I) + (t.I + 1e-9 >= r.acc ? ', over the ' + pctNice(r.acc) + ' line.' : ', still under the line. These students only help lift the phase total.')
        : dn + ' fewer counted as better than expected (' + f.H + ' to ' + t.H + '), taking it from ' + pctNice(f.J) + ' to ' + pctNice(t.J) + '.';
      return '<li><span class="st-name">' + esc(gradeName(f.grade)) + '</span><span class="st-txt">' + esc(txt) + (exact ? '' : ' <span class="muted">Includes students beyond the grade’s own minimum.</span>') + '</span><span class="st-j">' + (f.judg !== t.judg ? judgTag(f.judg) + '<span class="arrow" aria-hidden="true">→</span>' + judgTag(t.judg) : '<span class="muted">Stays ' + esc(t.judg) + '</span>') + '</span></li>';
    }).join('') + autoTotals.map(function (c) {
      return '<li class="auto"><span class="st-name">' + esc(c.to.label) + '</span><span class="st-txt">Changes on its own because its grades improved: ' + num(c.from.G) + ' to ' + num(c.to.G) + ' of ' + c.to.E + ' students, ' + pctNice(c.from.I) + ' to ' + pctNice(c.to.I) + '. Nobody edits this row.</span><span class="st-j">' + judgTag(c.from.judg) + '<span class="arrow" aria-hidden="true">→</span>' + judgTag(c.to.judg) + '</span></li>';
    }).join('') + '</ul>';
    var why = '';
    if (sim.options.length > 1) {
      why = i === 0
        ? '<p class="why">Listed first because it needs the fewest students' + (o.extras === 0 ? ', and every grade gets exactly its own minimum, with no spare students placed anywhere.' : '.') + '</p>'
        : '<p class="why">Same number of judgement changes as Option A. ' + (o.students > sim.options[0].students ? 'It needs ' + (o.students - sim.options[0].students) + ' more student' + (o.students - sim.options[0].students === 1 ? '' : 's') + ' than Option A.' : 'Same number of students as Option A,') + (o.extras ? ' ' + o.extras + ' of its students go beyond a grade’s own minimum so the phase total can cross the line.' : '') + '</p>';
    }
    var left = '<p class="left"><strong>Still Weak:</strong> ' + (stillWeak.length ? stillWeak.map(function (x) { return esc(x.label) + ' (' + pctNice(x.I) + ')'; }).join(', ') : 'none') + '. That is ' + d.counts['Weak'] + ' of ' + d.n + ' = ' + pctNice(d.counts['Weak'] / (d.n || 1)) + (Math.abs(d.counts['Weak'] / (d.n || 1) - r.capWeak) < 1e-9 ? ', exactly at the limit.' : '.') + '</p>';
    return lead + why + steps + left +
      '<h4 class="sub">The boxes after this pretend fix</h4>' + renderDist(counted(o.model.items, r.mode), r, { intro: false }) +
      '<h4 class="sub">Grade by grade, with the pretend numbers</h4><p class="muted small">Highlighted rows changed result. Under each changed number you can see how many students were added and what the real number was.</p>' + renderTable(o.model, r, { base: base });
  }
  function renderSimHeader(sim, r) {
    if (!sim.needed) return '<p class="lead-ok"><span class="dot ok" aria-hidden="true"></span>The actual results already meet both limits, so no simulation is needed.</p>';
    if (sim.failed) return '<p class="lead-bad"><span class="dot bad" aria-hidden="true"></span>' + (sim.tooMany ? 'There are too many Weak judgements to try every combination in the browser.' : 'No simulation can meet the limits by changing student counts alone. Try a different way of counting judgements in the rules.') + '</p>';
    var n = sim.options.length;
    return '<p class="sim-count">The fewest changes possible is <strong>' + sim.options[0].changed.length + ' judgements</strong>. ' + (n === 1 ? 'Only one way does it.' : 'There are ' + n + ' ways to do that. ' + (n > 6 ? 'The 6 that need the fewest students are shown' : 'They are shown') + ', starting with the fewest students.') + '</p>';
  }
  function renderTabs(sim, sel, prefix) {
    prefix = prefix || 'tab';
    if (!sim.options || sim.options.length < 2) return '';
    var show = Math.min(sim.options.length, 6);
    var panel = prefix === 'tab' ? 'opt-panel' : 'sheet-panel';
    var h = '<div class="tabs" role="tablist" aria-label="Simulation options" data-prefix="' + prefix + '">';
    for (var i = 0; i < show; i++) {
      var o = sim.options[i];
      h += '<button type="button" role="tab" id="' + prefix + '-' + i + '" aria-controls="' + panel + '" aria-selected="' + (i === sel) + '" tabindex="' + (i === sel ? 0 : -1) + '" data-opt="' + i + '"><span class="t-name">Option ' + letter(i) + '</span><span class="t-sub">' + o.students + ' more students' + (i === 0 ? ', fewest' : '') + '</span></button>';
    }
    h += '</div>';
    if (sim.options.length > show) h += '<p class="muted small">' + (sim.options.length - show) + ' more ways use the same number of judgement changes but need more students. They are not shown.</p>';
    return h;
  }

  // The loaded sheet exactly as received, in its own column layout, with phase labels merged as in the file.
  function renderSheetRaw(parsed) {
    if (!parsed || !parsed.raw || !parsed.raw.length) return '<p class="muted">No sheet loaded.</p>';
    var rowsBy = {}; parsed.rows.forEach(function (x) { rowsBy[x.grade] = x; });
    var raw = parsed.raw, spans = [], k = 0;
    while (k < raw.length) { var j = k; while (j + 1 < raw.length && raw[j + 1].label === raw[k].label) j++; spans.push([k, j - k + 1]); k = j + 1; }
    var spanAt = {}; spans.forEach(function (s) { spanAt[s[0]] = s[1]; });
    var hasErr = { lab: false, cnt: false, pct: false };
    var body = raw.map(function (x, i) {
      var isT = /total/i.test(x.grade), pr = rowsBy[x.grade], labErr = pr && pr.label !== pr.phase && parsed.totals.length;
      var mem = isT ? parsed.rows.filter(function (g) { return x.members && x.members.indexOf(g.grade) >= 0; }) : [];
      var sG = 0, sH = 0, sE = 0; mem.forEach(function (g) { sE += g.E; sG += g.G; sH += g.H; });
      var gErr = isT && x.G != null && Math.abs(x.G - sG) > 1e-6, hErr = isT && x.H != null && Math.abs(x.H - sH) > 1e-6;
      var iErr = isT && x.I != null && sE && Math.abs(x.I - sG / sE) > 5e-5, jErr = isT && x.J != null && sE && Math.abs(x.J - sH / sE) > 5e-5;
      if (gErr || hErr) hasErr.cnt = true; if (iErr || jErr) hasErr.pct = true;
      var labCell = '';
      if (spanAt[i] != null) labCell = '<td class="mg" rowspan="' + spanAt[i] + '">' + esc(x.label == null ? '' : x.label) + '</td>';
      var gradeErr = raw.slice(i, i + 1).some(function () { return labErr; });
      if (gradeErr) hasErr.lab = true;
      function c(v, bad, mark, fmt) { return '<td class="n' + (bad ? ' err' : '') + '">' + (fmt ? fmt(v) : num(v)) + (bad ? '<sup>' + mark + '</sup>' : '') + '</td>'; }
      return '<tr' + (isT ? ' class="tot"' : '') + '>' + labCell + '<th scope="row"' + (gradeErr ? ' class="err"' : '') + '>' + esc(x.grade) + (gradeErr ? '<sup>1</sup>' : '') + '</th>' +
        c(x.E) + (x.F != null ? c(x.F) : '<td></td>') + c(x.G, gErr, 2) + c(x.H, hErr, 2) + c(x.I, iErr, 3, function (v) { return pct(v); }) + c(x.J, jErr, 3, function (v) { return pct(v); }) +
        '<td>' + (x.K ? judgTag(x.K) : '') + '</td></tr>';
    }).join('');
    var notes = [];
    if (hasErr.lab) notes.push('<li value="1">This grade sits under the wrong phase label. The totals show it belongs to the next phase.</li>');
    if (hasErr.cnt) notes.push('<li value="2">The file averaged the grades here instead of adding them up.</li>');
    if (hasErr.pct) notes.push('<li value="3">This percentage treats every grade the same size, so it is slightly off.</li>');
    return '<div class="tbl-scroll"><table class="raw sheet"><thead><tr><th scope="col">Phase</th><th scope="col">Grade</th><th scope="col" class="n">Students</th><th scope="col" class="n">% assessed</th><th scope="col" class="n">Made expected or better</th><th scope="col" class="n">Did better than expected</th><th scope="col" class="n">% expected or better</th><th scope="col" class="n">% better than expected</th><th scope="col">Judgement</th></tr></thead><tbody>' + body + '</tbody></table></div>' +
      (notes.length ? '<ol class="notes">' + notes.join('') + '</ol>' : '');
  }

  // A simulated sheet laid out like the Your Data table: boxed values, changed boxes highlighted.
  function renderSimSheet(sim, i, r) {
    var o = sim.options[i], base = sim.model, lookup = {};
    base.items.forEach(function (it) { lookup[it.id] = it; });
    var body = '';
    o.model.phases.forEach(function (p) {
      var list = p.grades.slice(); if (p.total) list.push(p.total);
      list.forEach(function (it) {
        var b = lookup[it.id], isT = it.kind === 'total';
        function box(field) {
          var ch = b && b[field] !== it[field], d = ch ? it[field] - b[field] : 0;
          return '<td><span class="box' + (ch ? ' chg' : '') + (isT ? ' sum' : '') + '">' + num(it[field]) + '</span>' + (ch ? '<span class="box-note">' + (d > 0 ? '+' : '') + num(d) + ', was ' + num(b[field]) + '</span>' : '') + '</td>';
        }
        body += '<tr class="' + (isT ? 'tot' : '') + (b && b.judg !== it.judg ? ' moved' : '') + '">' +
          '<td><span class="box sum">' + esc(it.phase) + '</span></td>' +
          '<td><span class="box sum left">' + esc(isT ? 'Total' : it.grade) + '</span></td>' +
          box('E') + box('G') + box('H') +
          '<td class="jc">' + judgTag(it.judg) + (b && b.judg !== it.judg ? '<span class="was-j">was ' + esc(b.judg) + '</span>' : '') + '</td></tr>';
      });
    });
    return '<div class="tbl-scroll"><table class="edit simsheet"><thead><tr><th scope="col">Phase</th><th scope="col">Grade</th><th scope="col">Students</th><th scope="col">Made expected or better</th><th scope="col">Did better than expected</th><th scope="col" class="lft">Result</th></tr></thead><tbody>' + body + '</tbody></table></div>';
  }
  /* ---------- spreadsheet-style grid (mirrors the file's own layout and colours) ---------- */
  var XL = { yellow: '#FFFF00', cream: '#FEF2CB', pale: '#FEFFB3', orange: '#F6B26B', accent: '#ED7D31', red: '#FF0000' };
  var XL_HEAD = ['Phase', 'Group of Students', 'Grade/Year', 'No. of students in this group', '% Students assessed', 'students that made expected or better progress', 'students that made better than expected progress', '% of students that made expected or better progress', '% of students that made better than expected progress', 'Judgement'];
  function headRow() {
    return XL_HEAD.map(function (h, i) { return { v: h, t: 's', bold: true, size: 10, fill: i === 9 ? XL.yellow : null, wrap: true }; });
  }
  function bodyCells(o) {
    // o: {phase, grade, E, F, G, H, I, J, K, isT, marks:{G:'err'|'sim',...}, tips:{}}
    var red = o.isT ? XL.red : null, b = !!o.isT;
    function m(f) { return o.marks && o.marks[f]; }
    function tip(f) { return o.tips && o.tips[f]; }
    return [
      null, null,
      { v: o.grade, t: 's', fill: XL.cream, bold: b, color: red, mark: m('D'), tip: tip('D') },
      { v: o.E, t: 'n', fill: o.isT ? XL.cream : XL.yellow, bold: true, color: red, mark: m('E'), tip: tip('E') },
      { v: o.F == null ? 100 : o.F, t: 'n', fill: XL.cream },
      { v: o.G, t: 'n', fill: o.isT ? XL.cream : XL.yellow, bold: true, color: red, mark: m('G'), tip: tip('G') },
      { v: o.H, t: 'n', fill: o.isT ? XL.cream : XL.yellow, bold: true, color: red, mark: m('H'), tip: tip('H') },
      { v: o.I, t: 'p', fill: XL.cream, bold: b, color: red, mark: m('I'), tip: tip('I') },
      { v: o.J, t: 'p', fill: XL.cream, bold: b, color: red, mark: m('J'), tip: tip('J') },
      { v: o.K, t: 's', fill: o.isT ? XL.accent : XL.pale, mark: m('K'), tip: tip('K') }
    ];
  }
  function spanRuns(keys) {
    var out = {}, k = 0;
    while (k < keys.length) { var j = k; while (j + 1 < keys.length && keys[j + 1] === keys[k]) j++; out[k] = j - k + 1; k = j + 1; }
    return out;
  }
  function sheetSpecRaw(parsed, title) {
    var rowsBy = {}; parsed.rows.forEach(function (x) { rowsBy[x.grade] = x; });
    var raw = parsed.raw, notes = {};
    var body = raw.map(function (x) {
      var isT = /total/i.test(x.grade), pr = rowsBy[x.grade], marks = {}, tips = {};
      if (pr && pr.label !== pr.phase && parsed.totals.length) { marks.D = 'err'; tips.D = 'Labelled Phase ' + pr.label + ', but counted in the Phase ' + pr.phase + ' total'; notes.lab = 1; }
      if (isT) {
        var mem = parsed.rows.filter(function (g) { return x.members && x.members.indexOf(g.grade) >= 0; });
        var sE = 0, sG = 0, sH = 0; mem.forEach(function (g) { sE += g.E; sG += g.G; sH += g.H; });
        if (x.G != null && Math.abs(x.G - sG) > 1e-6) { marks.G = 'err'; tips.G = 'Average, not total. Should be ' + sG; notes.cnt = 1; }
        if (x.H != null && Math.abs(x.H - sH) > 1e-6) { marks.H = 'err'; tips.H = 'Average, not total. Should be ' + sH; notes.cnt = 1; }
        if (x.I != null && sE && Math.abs(x.I - sG / sE) > 5e-5) { marks.I = 'err'; tips.I = 'Should be ' + pct(sG / sE, 2); notes.pct = 1; }
        if (x.J != null && sE && Math.abs(x.J - sH / sE) > 5e-5) { marks.J = 'err'; tips.J = 'Should be ' + pct(sH / sE, 2); notes.pct = 1; }
      }
      return bodyCells({ grade: x.grade, E: x.E, F: x.F, G: x.G, H: x.H, I: x.I, J: x.J, K: x.K, isT: isT, marks: marks, tips: tips });
    });
    var spans = spanRuns(raw.map(function (x) { return x.label; }));
    body.forEach(function (row, i) {
      row[0] = spans[i] != null ? { v: raw[i].label, t: 'n', fill: XL.pale, rs: spans[i] } : 'skip';
      row[1] = i === 0 ? { v: 'Full Cohort', t: 's', fill: XL.orange, bold: true, size: 16, rot: true, rs: raw.length } : 'skip';
    });
    var n = [];
    if (notes.lab) n.push('Red corner on a grade: it sits under the wrong phase label. The totals show it belongs to the next phase.');
    if (notes.cnt) n.push('Red corner on a total count: the file averaged the grades instead of adding them up.');
    if (notes.pct) n.push('Red corner on a total percentage: it treats every grade as the same size, so it is slightly off.');
    return { title: title || 'Your sheet', head: headRow(), body: body, notes: n, pctDigits: 0, meta: raw.map(function (x) { return { isT: /total/i.test(x.grade) }; }) };
  }
  function sheetSpecModel(model, r, base, title) {
    var lookup = {}; if (base) base.items.forEach(function (it) { lookup[it.id] = it; });
    var body = [], phaseKeys = [], meta = [];
    model.phases.forEach(function (p) {
      var list = p.grades.slice(); if (p.total) list.push(p.total);
      list.forEach(function (it) {
        var b = lookup[it.id], marks = {}, tips = {};
        if (b) {
          if (b.G !== it.G) { marks.G = 'sim'; tips.G = 'Pretend. Real value ' + num(b.G) + ' (' + (it.G - b.G > 0 ? '+' : '') + num(it.G - b.G) + ')'; marks.I = 'sim'; tips.I = 'Real value ' + pct(b.I); }
          if (b.H !== it.H) { marks.H = 'sim'; tips.H = 'Pretend. Real value ' + num(b.H); marks.J = 'sim'; tips.J = 'Real value ' + pct(b.J); }
          if (b.judg !== it.judg) { marks.K = 'sim'; tips.K = 'Real result: ' + b.judg; }
        }
        body.push(bodyCells({ grade: it.kind === 'total' ? 'Total' : it.grade, E: it.E, G: it.G, H: it.H, I: it.I, J: it.J, K: it.judg, isT: it.kind === 'total', marks: marks, tips: tips, }));
        phaseKeys.push(String(p.phase)); meta.push({ isT: it.kind === 'total', G: it.G, baseG: b ? b.G : null });
      });
    });
    var spans = spanRuns(phaseKeys);
    body.forEach(function (row, i) {
      row[0] = spans[i] != null ? { v: phaseKeys[i], t: 'n', fill: XL.pale, rs: spans[i] } : 'skip';
      row[1] = i === 0 ? { v: 'Full Cohort', t: 's', fill: XL.orange, bold: true, size: 16, rot: true, rs: body.length } : 'skip';
    });
    return { title: title || 'Your sheet', head: headRow(), body: body, notes: [], pctDigits: 1, meta: meta };
  }
  function colLetter(i) { return String.fromCharCode(65 + i); }
  function cellVal(c, pd) {
    if (c.v == null || c.v === '') return '';
    if (c.t === 'p') return (c.v * 100).toFixed(pd || 0) + '%';
    if (c.t === 'n') return num(c.v);
    return esc(c.v);
  }
  function renderGrid(spec, opts) {
    opts = opts || {};
    var cols = 11; // A..K
    var h = '<div class="xl-scroll" tabindex="0" role="region" aria-label="' + esc(opts.label || 'Spreadsheet') + '"><table class="xl-grid"><colgroup><col class="xl-rn"><col class="xl-a">';
    for (var c = 1; c < cols; c++) h += '<col class="xl-c">';
    h += '</colgroup><thead><tr><th class="xl-corner" aria-hidden="true"></th>';
    for (c = 0; c < cols; c++) h += '<th scope="col" class="xl-col">' + colLetter(c) + '</th>';
    h += '</tr></thead><tbody>';
    var rn = 1;
    h += '<tr class="xl-r1"><th scope="row" class="xl-rown">' + (rn++) + '</th>' + '<td></td>'.repeat(cols) + '</tr>';
    h += '<tr class="xl-r2"><th scope="row" class="xl-rown">' + (rn++) + '</th><td></td><td colspan="10" class="xl-title" style="background:' + XL.yellow + '">' + esc(spec.title) + '</td></tr>';
    h += '<tr class="xl-r3"><th scope="row" class="xl-rown">' + (rn++) + '</th><td></td>' + spec.head.map(function (cell, i) {
      return '<td class="xl-h" style="' + (cell.fill ? 'background:' + cell.fill + ';' : '') + 'border-width:2px ' + (i === 0 || i === 9 ? 2 : 1) + 'px 2px ' + (i === 0 ? 2 : 1) + 'px">' + esc(cell.v) + '</td>';
    }).join('') + '</tr>';
    var nRows = spec.body.length;
    spec.body.forEach(function (row, ri) {
      var m = (spec.meta && spec.meta[ri]) || {}, nextT = spec.meta && spec.meta[ri + 1] && spec.meta[ri + 1].isT;
      h += '<tr><th scope="row" class="xl-rown">' + (rn++) + '</th><td></td>';
      row.forEach(function (cell, ci) {
        if (cell === 'skip' || cell == null) return;
        var cls = ['xl-v'];
        // borders as in the file: thin inside, medium round the block, the phase column and the total rows
        var bt = (ri === 0 || m.isT) ? 2 : 1, bb = (m.isT || ri === nRows - 1 || (cell.rs > 1 && ri + cell.rs >= nRows)) ? 2 : 1;
        if (cell.rs > 1) bb = 2;
        if (ci === 0) bt = 2;
        var bl = (ci === 0 || ci === 2) ? 2 : 1, br = (ci === 0 || ci === 1 || ci === 9) ? 2 : 1;
        if (nextT && ci > 1) bb = 2;
        if (cell.t === 'n' || cell.t === 'p') cls.push('xl-n');
        if (cell.bold) cls.push('xl-b');
        if (cell.rot) cls.push('xl-rot');
        if (cell.mark === 'err') cls.push('xl-err');
        if (cell.mark === 'sim') cls.push('xl-sim');
        var st = ['border-width:' + bt + 'px ' + br + 'px ' + bb + 'px ' + bl + 'px'];
        if (cell.fill && cell.mark !== 'sim') st.push('background:' + cell.fill);
        if (cell.color && cell.mark !== 'sim') st.push('color:' + cell.color);
        if (cell.size) st.push('font-size:' + (cell.size >= 16 ? '15px' : '12px'));
        h += '<td class="' + cls.join(' ') + '"' + (cell.rs > 1 ? ' rowspan="' + cell.rs + '"' : '') + (st.length ? ' style="' + st.join(';') + '"' : '') + (cell.tip ? ' title="' + esc(cell.tip) + '"' : '') + '>' +
          (cell.rot ? '<span>' + cellVal(cell, spec.pctDigits) + '</span>' : cellVal(cell, spec.pctDigits)) + (cell.tip ? '<span class="sr"> (' + esc(cell.tip) + ')</span>' : '') + '</td>';
      });
      h += '</tr>';
    });
    h += '</tbody></table></div>';
    return h;
  }
  function renderSheetTabs(names, sel, prefix, panel) {
    return '<div class="xl-tabs" role="tablist" aria-label="Sheets" data-prefix="' + prefix + '">' + names.map(function (n, i) {
      return '<button type="button" role="tab" id="' + prefix + '-' + i + '" aria-controls="' + panel + '" aria-selected="' + (i === sel) + '" tabindex="' + (i === sel ? 0 : -1) + '" data-opt="' + i + '">' + esc(n) + '</button>';
    }).join('') + '</div>';
  }
  function renderChangeList(sim, i) {
    var o = sim.options[i];
    var g = o.cellChanges.filter(function (c) { return c.from.kind === 'grade'; });
    var t = o.changed.filter(function (c) { return c.from.kind === 'total'; });
    return '<ul class="chg-list">' + g.map(function (c) {
      return '<li><strong>' + esc(gradeName(c.to.grade)) + '</strong> ' + c.from.G + ' to ' + c.to.G + ' of ' + c.to.E + ' <span class="chg-d">+' + (c.to.G - c.from.G) + '</span> ' + (c.from.judg !== c.to.judg ? judgTag(c.from.judg) + '<span class="arrow" aria-hidden="true">→</span>' + judgTag(c.to.judg) : '') + '</li>';
    }).join('') + t.map(function (c) {
      return '<li><strong>Phase ' + esc(c.to.phase) + ' total</strong> changes on its own ' + judgTag(c.from.judg) + '<span class="arrow" aria-hidden="true">→</span>' + judgTag(c.to.judg) + '</li>';
    }).join('') + '</ul>';
  }
  // Rows for an unstyled export when the data has been edited (SheetJS).
  function specToAoa(spec) {
    var aoa = [[], [null, spec.title], [null].concat(spec.head.map(function (c) { return c.v; }))], merges = [{ s: { r: 1, c: 1 }, e: { r: 1, c: 10 } }];
    spec.body.forEach(function (row, i) {
      var line = [null];
      row.forEach(function (cell, ci) {
        if (cell === 'skip' || cell == null) { line.push(null); return; }
        line.push(cell.v);
        if (cell.rs > 1) merges.push({ s: { r: 3 + i, c: 1 + ci }, e: { r: 3 + i + cell.rs - 1, c: 1 + ci } });
      });
      aoa.push(line);
    });
    return { aoa: aoa, merges: merges };
  }
  return {
    CATS: CATS, DEFAULT_RULES: DEFAULT_RULES, esc: esc, pct: pct, num: num, judge: judge, gradeName: gradeName, needAt: needAt,
    parseGrid: parseGrid, parseText: parseText, build: build, counted: counted, dist: dist, meets: meets, simulate: simulate,
    renderDist: renderDist, renderTable: renderTable, renderFixes: renderFixes, findFixes: findFixes, renderCompare: renderCompare,
    renderOption: renderOption, renderSimHeader: renderSimHeader, renderTabs: renderTabs, judgTag: judgTag, letter: letter, renderSheetRaw: renderSheetRaw, renderSimSheet: renderSimSheet, sheetSpecRaw: sheetSpecRaw, sheetSpecModel: sheetSpecModel, renderGrid: renderGrid, renderSheetTabs: renderSheetTabs, renderChangeList: renderChangeList, specToAoa: specToAoa
  };
})();
if (typeof module !== 'undefined') module.exports = Core;
