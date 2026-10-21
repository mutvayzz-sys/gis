  /* ---------- the designed PDF report (same layout as the hand-made English report) ---------- */
  // Builds one self-contained HTML document from what is loaded. The site prints it to PDF with
  // Cloudflare Browser Run; elsewhere the browser's own print dialog saves it as PDF.
  function letter(i) { return String.fromCharCode(65 + i); }
  function P1(v) { return v == null ? '–' : (Math.round(v * 1000) / 10) + '%'; }
  function sec(id, num, title, inner, wide) {
    return '<section class="pg' + (wide ? ' wide' : '') + '" id="r-' + id + '"><h2>' + (num ? '<span class="num">' + num + '</span>' : '') + esc(title) + '</h2>' + inner + '</section>';
  }
  function simBanner() { return '<p class="simban">SIMULATED. Not actual results. Do not report as findings.</p>'; }
  // Sheets wider than the page are scaled down to fit (landscape A4 is about 1000 CSS pixels wide).
  function gridOf(wb, name, maxCols, room) {
    var ws = wb && wb.getWorksheet(name); if (!ws) return '';
    var cols = Math.min(maxCols || 26, ws.columnCount || 26), w = 26;
    for (var c = 1; c <= cols; c++) w += Math.max(18, Math.min(220, (ws.getColumn(c).width || 9) * 7.2));
    var z = Math.min(1, (room || 1000) / w);
    return '<div class="xl"><div class="xl-scroll"' + (z < 1 ? ' style="zoom:' + z.toFixed(3) + '"' : '') + '>' + Grid.render(ws, { maxRows: 120, maxCols: maxCols || 26 }) + '</div></div>';
  }
  function judgementsTitle(subj) { return /progress$/i.test(subj) ? subj + ' Judgements' : subj + ' Progress Judgements'; }
  function tableCols() { var mx = 0; st.parsed.blocks.forEach(function (b) { Object.keys(b.col).forEach(function (k) { if (b.col[k] != null) mx = Math.max(mx, b.col[k]); }); }); return mx || 26; }
  function gradeRange(items) {
    var g = items.filter(function (x) { return x.kind === 'grade'; });
    if (!g.length) return '';
    return g.length === 1 ? g[0].grade : g[0].grade + ' to ' + g[g.length - 1].grade;
  }
  function reportHtml() {
    var c = st.cfg, m = st.base, s = st.sim, multi = st.parsed.blocks.length > 1;
    var k = E.check(E.counted(m.items, c.basis), c), fx = E.findFixes(st.parsed, c).fixes;
    var opts = (s && s.needed && !s.failed && s.options) ? s.options.slice(0, 6) : [];
    var sel = opts.length ? Math.min(st.sel, opts.length - 1) : 0;
    var subj = subjectName(), today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    var nG = m.items.filter(function (x) { return x.kind === 'grade'; }).length, nT = m.items.length - nG;
    var basisTxt = { both: nG + ' grades and ' + nT + ' phase totals', grades: nG + ' grades', phases: nT + ' phase totals' }[c.basis];
    var groups = Array.from(new Set(m.items.map(function (x) { return x.group; })));
    var label = function (it) { return (multi ? it.block + ': ' : '') + (it.kind === 'total' ? it.label : it.grade); };
    // Workbooks: one for the actual results and the selected option, one per other option for its sheet.
    var wbSel = Book.build({ srcWb: st.src, origWb: st.origWb, sheetName: st.sheetName, parsed: st.parsed, cfg: c, base: m, sim: s, sel: sel });
    var simName = Book.simSheetName(c);
    // How many file judgements the corrections changed
    var fileJ = []; st.parsed.blocks.forEach(function (b) { b.rows.forEach(function (x) { if (x.K) fileJ.push({ id: b.id + ':' + x.r, K: x.K }); }); });
    var byId = {}; m.items.forEach(function (x) { byId[x.id] = x; });
    var jChanged = fileJ.filter(function (f) { return byId[f.id] && byId[f.id].judg !== f.K; }).length;
    var realFixes = fx.filter(function (f) { return ['formulas', 'phase', 'derived'].indexOf(f.kind) < 0 && !/no formulas/i.test(f.detail); });
    var noPhase = fx.filter(function (f) { return f.kind === 'phase'; }), derived = fx.some(function (f) { return f.kind === 'derived'; });
    var A = opts[sel];

    /* cover */
    var findings = [];
    if (realFixes.length) findings.push('<strong>The file had ' + realFixes.length + ' adding-up or layout mistake' + (realFixes.length === 1 ? '' : 's') + (jChanged ? ', and fixing them changes ' + jChanged + ' judgement' + (jChanged === 1 ? '' : 's') : ', but fixing them changes no judgement') + '.</strong> ' + esc(realFixes.map(function (f) { return f.what; }).join('; ')) + '.');
    else findings.push('<strong>No adding-up mistakes were found.</strong> Every count, percentage and total checked out.');
    if (noPhase.length) findings.push('<strong>The sheet does not say which phase each grade belongs to.</strong> ' + (noPhase.length > 1 || multi ? 'Each table is' : 'All grades are') + ' treated as one phase for now, so the phase results need checking.');
    if (derived) findings.push('<strong>Only percentages were given.</strong> Student counts were worked out from them and rounded to whole students.');
    findings.push('<strong>The actual results ' + (k.weakOk ? 'meet' : 'do not meet') + ' the Weak limit.</strong> ' + k.weak + ' of ' + k.n + ' judgements are Weak (' + P1(k.n ? k.weak / k.n : 0) + '). The limit allows ' + k.maxW + '. ' + (k.outOk ? 'Outstanding is within its limit, with ' + k.out + '.' : 'Outstanding is over its limit: ' + k.out + ' against ' + k.maxO + ' allowed.'));
    k.extra.forEach(function (x) { findings.push('<strong>' + esc(x.label) + ': ' + (x.ok ? 'met' : 'not met') + '.</strong> ' + x.count + ' of ' + k.n + ' judgements.'); });
    if (A) findings.push('<strong>At least ' + (k.weak - k.maxW > 0 ? k.weak - k.maxW + ' judgements would have to stop being Weak' : 'a few judgements would have to change') + '.</strong> The smallest way found is ' + A.students + ' more students across ' + A.cells.length + ' grade' + (A.cells.length === 1 ? '' : 's') + ', changing ' + A.changed.length + ' judgements. ' + s.total + ' different way' + (s.total === 1 ? '' : 's') + ' reach the limit with that many changes.' + (A.substantial ? ' The largest change in one grade is ' + A.maxPts.toFixed(1) + ' percentage points, which is substantial.' : ''));
    else if (s && s.needed && s.failed) findings.push('<strong>No simulated scenario can meet the limits with these rules.</strong>');
    if (A) findings.push('<strong>The simulations are hypothetical.</strong> They show the size of improvement needed. They are not actual results and are marked as simulated on every page.');
    if (m.items.some(function (x) { return x.roundingFlag; })) findings.push('<strong>A rounding rule is needed.</strong> Some phase averages have decimals and no rounding rule was given, so the nearest whole point is used for now.');
    var toc = ['How judgements are decided', 'The sheet as received', 'Corrections made', 'Corrected results', 'Judgement counts'].concat(A ? ['Simulated options' + (opts.length > 1 ? ' A to ' + letter(opts.length - 1) : '')] : []).concat(['Notes']);
    var cover = '<section class="cover">' +
      '<p class="eyebrow">Progress review</p><h1>' + esc(judgementsTitle(subj)) + '</h1>' +
      '<p class="sub">Checking the overall judgements against the limits: no more than ' + P1(c.capWeak) + ' Weak and no more than ' + P1(c.capOut) + ' Outstanding.</p>' +
      '<dl class="meta"><div><dt>Prepared</dt><dd>' + esc(today) + '</dd></div><div><dt>Source file</dt><dd>' + esc(st.srcName || '') + '</dd></div><div><dt>Group</dt><dd>' + esc(groups.join(', ') + (gradeRange(m.items) ? ', ' + gradeRange(m.items) : '')) + '</dd></div></dl>' +
      '<dl class="stats rep-stats">' +
        '<div><dt>Judgements counted</dt><dd>' + k.n + '<small>' + esc(basisTxt) + '</small></dd></div>' +
        '<div><dt>Weak now</dt><dd class="' + (k.weakOk ? '' : 'bad') + '">' + k.weak + ' of ' + k.n + '<small>' + P1(k.n ? k.weak / k.n : 0) + (k.weakOk ? ', within the limit' : ', over the limit') + '</small></dd></div>' +
        '<div><dt>Weak allowed</dt><dd>' + k.maxW + '<small>' + P1(c.capWeak) + ' of ' + k.n + '</small></dd></div>' +
        '<div><dt>Smallest simulated fix</dt><dd class="sim">' + (A ? A.changed.length + ' changes<small>' + A.students + ' more students, simulated</small>' : (s && s.needed ? 'None found<small>with these rules</small>' : 'Not needed<small>targets already met</small>')) + '</dd></div>' +
      '</dl>' +
      '<h3>Key findings</h3><ol class="findings">' + findings.map(function (f) { return '<li>' + f + '</li>'; }).join('') + '</ol>' +
      '<h3>Contents</h3><ol class="toc">' + toc.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('') + '</ol></section>';

    /* 1 criteria */
    var crit = [];
    if (c.vw != null) crit.push(['Very Weak', 'Fewer than this share make expected or better progress', 'below ' + P1(c.vw)]);
    crit.push(['Weak', 'Fewer than this share make expected or better progress', 'below ' + P1(c.acc)], ['Acceptable', 'At least this share make expected or better progress', P1(c.acc)], ['Good', 'Acceptable, and at least this share did better than expected', P1(c.good)], ['Very Good', 'Acceptable, and at least this share did better than expected', P1(c.vg)], ['Outstanding', 'Acceptable, and at least this share did better than expected', P1(c.out)]);
    var limits = [['Weak', 'No more than ' + P1(c.capWeak) + ' of all overall judgements may be Weak' + (c.vwCountsAsWeak && c.vw != null ? ' (Very Weak counts as Weak)' : ''), k.maxW + ' of ' + k.n], ['Outstanding', 'No more than ' + P1(c.capOut) + ' of all overall judgements may be Outstanding', k.maxO + ' of ' + k.n]];
    var critHtml = '<p class="lead">Progress is judged on the share of students who made expected or better progress, and then on the share who did better than expected. Thresholds mean “at least”.</p>' +
      '<table class="cmp full"><thead><tr><th>Judgement</th><th>What is checked</th><th class="n">Line</th></tr></thead><tbody>' + crit.map(function (x) { return '<tr><th scope="row">' + tag(x[0]) + '</th><td>' + x[1] + '</td><td class="n">' + x[2] + '</td></tr>'; }).join('') + '</tbody></table>' +
      '<h3>Limits for the whole set</h3><table class="cmp full"><thead><tr><th>Limit</th><th>Rule</th><th class="n">Most allowed</th></tr></thead><tbody>' + limits.map(function (x) { return '<tr><th scope="row">' + tag(x[0]) + '</th><td>' + x[1] + '</td><td class="n">' + x[2] + '</td></tr>'; }).join('') +
      (c.extraRules || []).map(function (r, i) { var x = k.extra[i]; return '<tr><th scope="row">Extra rule</th><td>' + esc(r.label || E.ruleLabel(r)) + '</td><td class="n">' + (x ? (r.op === 'min' ? 'at least ' : 'at most ') + x.limit + ' of ' + k.n : '') + '</td></tr>'; }).join('') + '</tbody></table>' +
      '<p class="note">Counted toward the limits: ' + esc(basisTxt) + '. Phase totals are judged ' + (c.totalMethod === 'points' ? 'from the average of their grade judgements (points ' + E.CATS.filter(function (x) { return c.points[x] != null; }).slice().reverse().map(function (x) { return x + ' ' + c.points[x]; }).join(', ') + '), rounded ' + (c.rounding === 'unspecified' ? 'to the nearest point until a rule is confirmed' : c.rounding) : 'from their summed student counts') + '.' + (c.useAssessed ? ' Percentages are out of the students assessed.' : '') + '</p>';

    /* 2 original, 3 corrections, 4 corrected */
    var origHtml = '<p class="lead">Exactly as received. The mistakes are explained in section 3.</p>' + gridOf(wbSel, Book.NAMES.orig, null, 1000) +
      (realFixes.length ? '<ul class="xl-notes">' + realFixes.map(function (f) { return '<li>' + esc(f.what) + '.</li>'; }).join('') + '</ul>' : '');
    var corrHtml = '<p class="lead">' + (realFixes.length ? 'Only calculation and layout mistakes were fixed. No student numbers were changed.' : 'No mistakes needed fixing.') + (fx.length > realFixes.length ? ' Notes about how the sheet was read are listed too.' : '') + '</p>' +
      (fx.length ? '<table class="cmp full corr"><thead><tr><th class="n">#</th><th>Where</th><th>What was wrong</th><th>Effect on judgement</th></tr></thead><tbody>' + fx.map(function (f, i) { return '<tr><td class="n">' + (i + 1) + '</td><th scope="row">' + esc(f.what) + '</th><td>' + esc(f.detail) + '</td><td>' + esc(f.effect || '') + '</td></tr>'; }).join('') + '</tbody></table>' : '');
    var fixedHtml = '<p class="lead">The same sheet with the totals added up properly' + (fx.some(function (f) { return f.kind === 'label'; }) ? ' and the phase labels in the right place' : '') + '. These are the real results. Orange cells were corrected.</p>' + gridOf(wbSel, Book.NAMES.actual, multi ? tableCols() : null, 1000);
    var cards = phaseCards(m, c);
    var detailHtml = '<h3>The big picture</h3><p class="note">Each box is one judgement.</p>' + tilesBlock(m, null, c) +
      '<h3>Grade by grade</h3><p class="note">The bar is the share making expected or better progress; the black line is the ' + P1(c.acc) + ' mark.</p>' + barsBlock(m, c) +
      (cards.length ? '<div class="keep"><h3>Phase totals</h3><div class="phasecards">' + cards.join('') + '</div></div>' : '');

    /* 5 counts */
    var srcs = [['Corrected', E.counted(m.items, c.basis)]].concat(opts.map(function (o, i) { return ['Option ' + letter(i), E.counted(o.model.items, c.basis)]; }));
    var cats = E.CATS.slice().reverse().filter(function (x) { return x !== 'Very Weak' || c.vw != null; });
    var ct = '<table class="cmp full"><thead><tr><th>Judgement</th>' + srcs.map(function (x, i) { return '<th class="n' + (i ? ' simh' : '') + '">' + x[0] + '</th>'; }).join('') + '</tr></thead><tbody>';
    var chk = srcs.map(function (x) { return E.check(x[1], c); });
    cats.forEach(function (cat) { ct += '<tr><th scope="row">' + tag(cat) + '</th>' + chk.map(function (q) { return '<td class="n">' + q.d.counts[cat] + '</td>'; }).join('') + '</tr>'; });
    ct += '<tr class="tot"><th scope="row">Total</th>' + chk.map(function (q) { return '<td class="n">' + q.n + '</td>'; }).join('') + '</tr>';
    ct += '<tr><th scope="row">Weak share</th>' + chk.map(function (q) { return '<td class="n">' + P1(q.n ? q.weak / q.n : 0) + '</td>'; }).join('') + '</tr>';
    ct += '<tr><th scope="row">Within ' + P1(c.capWeak) + ' Weak?</th>' + chk.map(function (q) { return '<td class="n"><span class="status ' + (q.weakOk ? 'ok' : 'bad') + '">' + (q.weakOk ? 'Yes' : 'No') + '</span></td>'; }).join('') + '</tr>';
    ct += '<tr><th scope="row">Within ' + P1(c.capOut) + ' Outstanding?</th>' + chk.map(function (q) { return '<td class="n"><span class="status ' + (q.outOk ? 'ok' : 'bad') + '">' + (q.outOk ? 'Yes' : 'No') + '</span></td>'; }).join('') + '</tr>';
    (c.extraRules || []).forEach(function (r, ri) { ct += '<tr><th scope="row">' + esc(r.label || E.ruleLabel(r)) + '?</th>' + chk.map(function (q) { var x = q.extra[ri]; return '<td class="n"><span class="status ' + (x && x.ok ? 'ok' : 'bad') + '">' + (x && x.ok ? 'Yes' : 'No') + '</span></td>'; }).join('') + '</tr>'; });
    ct += '</tbody></table>';
    var diff = '';
    if (nG && nT) {
      var rowsD = [['The ' + nG + ' grades', 'grades'], ['The ' + nT + ' phase totals', 'phases']].concat(c.basis !== 'both' ? [['All ' + m.items.length + ' judgements', 'both']] : []).filter(function (x) { return x[1] !== c.basis; });
      diff = '<h3>If the judgements were counted differently</h3><table class="cmp full"><thead><tr><th>Counting only</th><th class="n">Corrected</th>' + (A ? '<th class="n simh">Option ' + letter(sel) + '</th>' : '') + '<th>Meaning</th></tr></thead><tbody>' + rowsD.map(function (x) {
        var qa = E.check(E.counted(m.items, x[1]), c), qs = A ? E.check(E.counted(A.model.items, x[1]), c) : null;
        var mean = (qa.weakOk ? 'The actual results would meet the Weak limit.' : 'The actual results would not meet the Weak limit (' + qa.maxW + ' allowed).') + (qs ? (qs.weakOk ? ' Option ' + letter(sel) + ' would also meet it.' : ' Option ' + letter(sel) + ' would not be enough.') : '');
        return '<tr><th scope="row">' + x[0] + '</th><td class="n">' + qa.weak + ' of ' + qa.n + ' Weak</td>' + (qs ? '<td class="n">' + qs.weak + ' of ' + qs.n + ' Weak</td>' : '') + '<td>' + mean + '</td></tr>';
      }).join('') + '</tbody></table>';
    }
    var countsHtml = '<p class="lead">How many of each judgement, out of ' + k.n + ', after correction' + (opts.length ? ' and in each simulated option. Options are simulated.' : '.') + '</p>' + ct + diff;

    /* 6 simulations */
    var simHtml = '', optPages = '';
    if (A) {
      var g0 = A.cells[0];
      var example = g0 ? '<p><strong>Example.</strong> ' + esc(label(g0.from)) + ' has ' + g0.from.A + ' assessed students and ' + g0.from.G + ' made expected or better progress (' + P1(g0.from.I) + '). Reaching ' + P1(c.acc) + ' needs ' + E.needAt(g0.from.A, c.acc) + ', so it is ' + Math.max(0, E.needAt(g0.from.A, c.acc) - g0.from.G) + ' short.' + (g0.dG ? ' With ' + g0.dG + ' more, it becomes ' + g0.to.judg + '.' : '') + '</p>' : '';
      var overview = '<table class="cmp full"><thead><tr><th>Option</th><th class="n">Students</th><th class="n">Changes</th><th>Grades changed</th><th>Still Weak</th></tr></thead><tbody>' + opts.map(function (o, i) {
        var gr = o.cells.map(function (x) { return label(x.to) + ' ' + (x.dG ? (x.dG > 0 ? '+' : '') + x.dG : '') + (x.dH ? ' (' + (x.dH > 0 ? '+' : '') + x.dH + ' better)' : ''); }).join(', ');
        var still = E.counted(o.model.items, c.basis).filter(function (x) { return E.isWeakish(x.judg, c); }).map(label).join(', ');
        return '<tr' + (i === sel ? ' class="selrow"' : '') + '><th scope="row">Option ' + letter(i) + (i === sel ? ' (chosen)' : '') + '</th><td class="n">' + o.students + '</td><td class="n">' + o.changed.length + '</td><td>' + esc(gr) + '</td><td>' + esc(still || 'None') + '</td></tr>';
      }).join('') + '</tbody></table>';
      var limitsTxt = [];
      if (c.simMaxStudents != null) limitsTxt.push('no grade gains more than ' + c.simMaxStudents + ' students');
      if (c.simMaxPts != null) limitsTxt.push('no grade moves more than ' + c.simMaxPts + ' percentage points');
      if ((c.simLocked || []).length) limitsTxt.push((c.simLocked || []).map(function (id) { return byId[id] ? label(byId[id]) : id; }).join(', ') + ' left alone');
      simHtml = simBanner() + '<p class="lead">To meet the limits, ' + (k.weak - k.maxW > 0 ? 'at least ' + (k.weak - k.maxW) + ' judgements must stop being Weak' : 'some judgements must change') + '. A grade stops being Weak when at least ' + P1(c.acc) + ' of its assessed students make expected or better progress. Each option assumes a few more students did, in as few grades as possible. Phase totals are never edited: they change on their own when their grades do.</p>' + example +
        '<p>' + s.total + ' different way' + (s.total === 1 ? '' : 's') + ' reach the limit with the minimum of ' + A.changed.length + ' changes.' + (opts.length > 1 ? ' The ' + opts.length + ' needing the fewest students are shown.' : '') + (limitsTxt.length ? ' Limits set for the simulation: ' + esc(limitsTxt.join('; ')) + '.' : '') + (s.pinned ? ' Some changes were chosen by hand and are included.' : '') + '</p>' + overview;
      optPages = opts.map(function (o, i) {
        var wbO = i === sel ? wbSel : Book.build({ srcWb: st.src, origWb: st.origWb, sheetName: st.sheetName, parsed: st.parsed, cfg: c, base: m, sim: s, sel: i });
        var list = '<ul class="chg-list">' + o.cells.map(function (x) { return '<li><b>' + esc(label(x.to)) + '</b> ' + x.from.G + ' to ' + x.to.G + ' of ' + x.from.A + (x.dH ? ', better than expected ' + x.from.H + ' to ' + x.to.H : '') + ': ' + tag(x.from.judg) + ' <span class="arrow">→</span> ' + tag(x.to.judg) + (x.pp > c.substantialPts ? ' <b class="bad">' + x.pp.toFixed(1) + ' points, substantial</b>' : '') + '</li>'; }).join('') +
          o.changed.filter(function (x) { return x.from.kind === 'total'; }).map(function (x) { return '<li><b>' + esc(label(x.to)) + '</b> changes on its own: ' + tag(x.from.judg) + ' <span class="arrow">→</span> ' + tag(x.to.judg) + '</li>'; }).join('') + '</ul>';
        // Long change lists go under the sheet in columns, so the page does not overflow.
        var stacked = o.cells.length + o.changed.filter(function (x) { return x.from.kind === 'total'; }).length > 8;
        return sec('opt' + i, '', 'Option ' + letter(i) + ': ' + o.students + ' more students' + (i === sel ? ' (chosen)' : ''), simBanner() + '<div class="optwrap' + (stacked ? ' stacked' : '') + '">' + gridOf(wbO, simName, tableCols(), stacked ? 1000 : 720) + '<div class="optside"><p class="small"><strong>' + o.students + ' more students</strong> make the progress shown. ' + o.changed.length + ' judgements change:</p>' + list + '<p class="note">Blue cells are simulated. All other cells are the real, corrected numbers.</p></div></div>', true);
      }).join('');
    }

    /* 7 notes */
    var notes = [];
    var border = [];
    m.items.filter(function (x) { return x.kind === 'grade' && x.A > 0; }).forEach(function (g) {
      [['acc', 'G', 'making expected or better progress', 'Acceptable'], ['good', 'H', 'doing better than expected', 'Good'], ['vg', 'H', 'doing better than expected', 'Very Good'], ['out', 'H', 'doing better than expected', 'Outstanding']].forEach(function (t) {
        if (g[t[1]] === E.needAt(g.A, c[t[0]]) && g.judg === t[3]) border.push(label(g) + ' has exactly ' + P1(g[t[1]] / g.A) + ' ' + t[2] + ' (' + g[t[1]] + ' of ' + g.A + '), so it is ' + t[3] + ' with no margin');
      });
    });
    if (border.length) notes.push('<strong>Borderline values.</strong> ' + esc(border.join('. ')) + '.');
    if (A) {
      var ks = A.check;
      if (ks.weak === ks.maxW) notes.push('<strong>Exactly at the limit.</strong> Option ' + letter(sel) + ' leaves ' + ks.weak + ' of ' + ks.n + ' Weak, the most the limit allows. The rule says “no more than”, so this meets it.');
      notes.push('<strong>How the options were found.</strong> Every combination of grades that could stop being Weak was tested. Only the number of students making expected or better progress was changed; cohort sizes, assessed students and thresholds were not touched, and every judgement was recalculated from the criteria.');
      notes.push('<strong>Simulations are not findings.</strong> They show the scale of improvement that would be needed to meet the limit.');
    }
    if (m.items.some(function (x) { return x.roundingFlag; })) notes.push('<strong>Rounding.</strong> No rounding rule was given for phase averages with decimals, so the nearest whole point is used for now. Please confirm the rule.');
    if (fx.some(function (f) { return f.kind === 'derived'; })) notes.push('<strong>Counts worked out from percentages.</strong> Where only percentages were given, student counts were worked out as percentage times assessed students, rounded to whole students.');
    notes.push('<strong>Accompanying workbook.</strong> “' + esc(fileBase()) + ' - Actual and Simulated Results.xlsx” holds the same sheets with live formulas.');
    var nt = notesOf();
    var asstHtml = '';
    if (nt.changes.length || nt.transcript.length) {
      asstHtml = sec('assistant', '', 'Assistant notes', '<p class="lead">What the assistant was asked, what it said, and what it changed on the page. The results in this report already include these changes.</p>' +
        (nt.changes.length ? '<h3>Changes made</h3><ol class="plain">' + nt.changes.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ol>' : '') +
        (nt.transcript.length ? '<h3>Conversation</h3>' + nt.transcript.map(function (t) { return '<div class="turn ' + t.role + '"><b>' + (t.role === 'user' ? 'Asked' : 'Assistant') + '</b>' + (t.role === 'user' ? '<p>' + esc(String(t.text).slice(0, 4000)).replace(/\n/g, '<br>') + '</p>' : md(String(t.text).slice(0, 12000))) + '</div>'; }).join('') : ''), false);
    }

    var n = 0, body = cover +
      sec('criteria', ++n, 'How judgements are decided', critHtml) +
      sec('original', ++n, 'The sheet as received', origHtml, true) +
      sec('corrections', ++n, 'Corrections made', corrHtml) +
      sec('fixed', ++n, 'Corrected results', fixedHtml, true) +
      sec('fixed-detail', '', 'Corrected results in detail', detailHtml) +
      sec('counts', ++n, 'Judgement counts', countsHtml) +
      (A ? sec('simulations', ++n, 'Simulated options', simHtml) + optPages : '') +
      sec('notes', ++n, 'Notes', '<ul class="plain">' + notes.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ul>') + asstHtml;

    var pageCss = Array.prototype.map.call(document.querySelectorAll('style'), function (x) { return x.textContent; }).join('\n');
    var title = esc(judgementsTitle(subj)) + ': Review';
    var footer = esc(judgementsTitle(subj)) + ': review, ' + esc(today);
    return '<!doctype html><html lang="en" data-theme="light"><head><meta charset="utf-8"><title>' + title + '</title>' +
      '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,400..700&display=swap">' +
      '<style>' + pageCss + '\n' + REPORT_CSS.split('%FOOTER%').join(footer.replace(/"/g, '\\"')) + '</style></head><body class="report">' + body + '</body></html>';
  }
  var REPORT_CSS = [
    ':root { color-scheme: light; }',
    '@page { size: A4; margin: 16mm 15mm 18mm; @bottom-left { content: "%FOOTER%"; font: 8pt Inter, Arial, sans-serif; color: #6e6e73; } @bottom-right { content: "Page " counter(page) " of " counter(pages); font: 8pt Inter, Arial, sans-serif; color: #6e6e73; } }',
    '@page wide { size: A4 landscape; margin: 12mm 12mm 16mm; @bottom-left { content: "%FOOTER%"; font: 8pt Inter, Arial, sans-serif; color: #6e6e73; } @bottom-right { content: "Page " counter(page) " of " counter(pages); font: 8pt Inter, Arial, sans-serif; color: #6e6e73; } }',
    'html, body { background: #fff !important; }',
    'body.report { font-family: Inter, Arial, sans-serif; font-size: 10.5pt; line-height: 1.5; color: #1d1d1f; margin: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }',
    'body.report .pg, body.report .cover { break-before: page; display: grid; gap: 12px; grid-template-columns: minmax(0, 1fr); padding: 0; }',
    'body.report .cover { break-before: auto; gap: 14px; }',
    'body.report .wide { page: wide; }',
    '.eyebrow { font-size: 9pt; letter-spacing: .06em; text-transform: uppercase; color: #6e6e73; font-weight: 600; margin: 0; }',
    'body.report h1 { font-size: 30pt; letter-spacing: -0.03em; margin: 0; line-height: 1.1; }',
    '.cover .sub { font-size: 12.5pt; color: #6e6e73; max-width: 60ch; margin: 0; }',
    '.meta { display: flex; flex-wrap: wrap; gap: 6px 28px; margin: 0; font-size: 9.5pt; }',
    '.meta dt, .rep-stats dt { color: #6e6e73; font-size: 8pt; text-transform: uppercase; letter-spacing: .04em; font-weight: 600; }',
    '.meta dd { margin: 0; font-weight: 550; }',
    '.rep-stats { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 0; margin: 4px 0 0; border-top: 1px solid #e3e3e6; border-bottom: 1px solid #e3e3e6; }',
    '.rep-stats > div { padding: 10px 12px 12px 0; } .rep-stats > div + div { padding-left: 12px; border-left: 1px solid #e3e3e6; }',
    '.rep-stats dd { margin: 4px 0 0; font-size: 16pt; font-weight: 650; line-height: 1.15; }',
    '.rep-stats dd small { display: block; font-size: 8.5pt; font-weight: 400; color: #6e6e73; margin-top: 3px; }',
    '.rep-stats dd.bad { color: #c00000; } .rep-stats dd.sim { color: #0000da; }',
    '.findings { margin: 0; padding-left: 1.3em; display: grid; gap: 8px; }',
    '.toc { margin: 0; padding-left: 1.3em; columns: 2; }',
    'body.report h2 { font-size: 17pt; display: flex; gap: 10px; align-items: baseline; margin: 0; letter-spacing: -0.01em; }',
    'body.report h2 .num { color: #a1a1a6; font-weight: 600; font-size: 13pt; }',
    'body.report h3 { font-size: 12pt; margin: 6px 0 0; break-after: avoid; }',
    'body.report h3 + *, body.report h2 + * { break-before: avoid; }',
    '.lead { color: #6e6e73; max-width: 90ch; margin: 0; }',
    '.note, .xl-notes { font-size: 9pt; color: #6e6e73; margin: 0; }',
    '.simban { background: #c00000; color: #fff; font-weight: 700; padding: 6px 12px; border-radius: 4px; font-size: 10pt; margin: 0; }',
    'table.cmp { border-collapse: collapse; width: 100%; font-size: 9.5pt; }',
    'table.cmp th, table.cmp td { text-align: left; padding: 7px 8px; border-bottom: 1px solid #e3e3e6; vertical-align: top; }',
    'table.cmp thead th { font-size: 8.5pt; color: #6e6e73; font-weight: 600; border-bottom: 1px solid #c7c7cc; }',
    'table.cmp tbody th { font-weight: 600; }',
    'table.cmp .n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }',
    'table.cmp .simh { color: #0000da; } table.cmp tr.tot th, table.cmp tr.tot td { border-top: 1px solid #c7c7cc; font-weight: 650; }',
    'table.cmp tr.selrow th, table.cmp tr.selrow td { background: #f4f4fe; }',
    '.status { font-weight: 650; } .status.ok { color: #1a7f37; } .status.bad, .bad { color: #c00000; }',
    '.corr td, .corr th { font-size: 9pt; }',
    'body.report .xl { border: 0; box-shadow: none; } body.report .xl-scroll { max-height: none; overflow: visible; }',
    'body.report .xl-grid { font-family: Calibri, Carlito, Arial, sans-serif; font-size: 9.5px; } body.report .xl-grid th, body.report .xl-grid td { padding: 1px 4px; } body.report .xl-grid thead th, body.report .xl-grid th.xl-rown { position: static; }',
    'body.report .xl-note { position: relative; } body.report .xl-note::after { content: ""; position: absolute; top: 0; right: 0; border-left: 6px solid transparent; border-top: 6px solid #d92b2b; }',
    'body.report .sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); }',
    '.optwrap { display: grid; grid-template-columns: auto minmax(220px, 1fr); gap: 16px; align-items: start; }',
    '.optside { display: grid; gap: 8px; } .chg-list { margin: 0; padding-left: 1.1em; display: grid; gap: 4px; font-size: 9pt; }',
    '.optwrap.stacked { grid-template-columns: minmax(0, 1fr); gap: 10px; } .optwrap.stacked .chg-list { display: block; columns: 3; column-gap: 20px; font-size: 8pt; } .optwrap.stacked .chg-list li { break-inside: avoid; margin-bottom: 3px; }',
    '.plain { margin: 0; padding-left: 1.2em; display: grid; gap: 8px; }',
    '.turn { border-top: 1px solid #e3e3e6; padding: 8px 0; } .turn b { display: block; font-size: 9pt; color: #6e6e73; } .turn.user b { color: #0000da; } .turn p { margin: 4px 0; } .turn ul { margin: 4px 0; }',
    'body.report .phasecards { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }',
    'body.report .pc, body.report .tiles-row, body.report .bar-row, table.cmp tr { break-inside: avoid; }',
    'body.report .keep { break-inside: avoid; display: grid; gap: 8px; }',
    'body.report .phasecards { break-inside: avoid; }',
    'section { break-inside: auto; }'
  ].join('\n');

  // PDF: on the site, Cloudflare prints it; otherwise the browser's print dialog saves it as PDF.
  function printReport(html) {
    return new Promise(function (resolve, reject) {
      var f = document.createElement('iframe');
      f.setAttribute('aria-hidden', 'true'); f.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0;';
      f.onload = function () {
        try { setTimeout(function () { f.contentWindow.focus(); f.contentWindow.print(); resolve('printed'); setTimeout(function () { f.remove(); }, 60000); }, 600); }
        catch (e) { f.remove(); reject(e); }
      };
      f.srcdoc = html; document.body.appendChild(f);
    });
  }
  function reportDownload(out) {
    if (!st.base) { msg(out, 'Load a sheet first.', 'bad'); return; }
    var html;
    try { html = reportHtml(); } catch (err) { msg(out, 'The report could not be built (' + err.message + ').', 'bad'); return; }
    var name = fileBase() + ' - Progress Report.pdf';
    var onSite = !(window.claude && typeof window.claude.use === 'function');
    msg(out, 'Making the PDF report…');
    var serverPdf = onSite ? fetch('/api/pdf', { method: 'POST', headers: { 'content-type': 'text/html' }, body: html }).then(function (r) {
      if (!r.ok) throw new Error('server ' + r.status);
      return r.arrayBuffer();
    }) : Promise.reject(new Error('no server'));
    return serverPdf.then(function (buf) { return saveFile(name, new Uint8Array(buf), out); }, function () {
      msg(out, 'Opening the print window. Choose “Save as PDF” as the printer.');
      return printReport(html).catch(function () { msg(out, 'Printing is blocked here, so a simpler PDF was made instead.', ''); return saveFile(fileBase() + ' - Report.pdf', reportPdf().then(function (b) { return new Uint8Array(b); }), out); });
    });
  }
