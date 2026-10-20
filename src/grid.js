/* Draws an ExcelJS worksheet as a spreadsheet-looking HTML table (values, merges, fills, fonts, borders). */
var Grid = (function () {
  var THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function color(c) {
    if (!c) return null;
    if (c.argb && c.argb.length === 8) return '#' + c.argb.slice(2);
    if (c.argb && c.argb.length === 6) return '#' + c.argb;
    if (c.theme != null && THEME[c.theme]) return '#' + THEME[c.theme];
    return null;
  }
  function display(cell) {
    var v = cell.value, fmt = cell.numFmt || '';
    if (v && typeof v === 'object') {
      if ('result' in v) v = v.result; else if (v.richText) v = v.richText.map(function (t) { return t.text; }).join('');
      else if ('text' in v) v = v.text; else if (v instanceof Date) v = v.toISOString().slice(0, 10); else if ('formula' in v || 'sharedFormula' in v) v = ''; else v = '';
    }
    if (v == null) return '';
    if (typeof v === 'number') {
      if (/%/.test(fmt)) { var m = /0\.(0+)%/.exec(fmt); return (v * 100).toFixed(m ? m[1].length : 0) + '%'; }
      if (/^0$/.test(fmt) || /^#,##0$/.test(fmt)) return String(Math.round(v));
      if (/0\.00/.test(fmt)) return v.toFixed(2);
      if (/\+0/.test(fmt)) return (v > 0 ? '+' : '') + Math.round(v);
      return Math.abs(v - Math.round(v)) < 1e-9 ? String(Math.round(v)) : String(Math.round(v * 100) / 100);
    }
    return String(v);
  }
  function bw(b) { if (!b || !b.style) return null; return b.style === 'thin' || b.style === 'hair' || b.style === 'dotted' ? 1 : 2; }
  function render(ws, opt) {
    opt = opt || {};
    var maxR = Math.min(ws.rowCount || 0, opt.maxRows || 200), maxC = Math.min(ws.columnCount || 0, opt.maxCols || 30);
    // trim empty trailing rows/cols
    var lastR = 0, lastC = 0;
    for (var r = 1; r <= maxR; r++) for (var c = 1; c <= maxC; c++) { var cl = ws.findCell ? ws.findCell(r, c) : ws.getCell(r, c); if (cl && cl.value != null && cl.value !== '') { lastR = Math.max(lastR, r); lastC = Math.max(lastC, c); } }
    maxR = Math.max(lastR, 1); maxC = Math.max(lastC, 1);
    var skip = {}, span = {};
    (ws.model.merges || []).forEach(function (m) {
      var mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m); if (!mm) return;
      var c1 = colNum(mm[1]), r1 = +mm[2], c2 = colNum(mm[3]), r2 = +mm[4];
      span[r1 + ',' + c1] = [r2 - r1 + 1, c2 - c1 + 1];
      for (var rr = r1; rr <= r2; rr++) for (var cc = c1; cc <= c2; cc++) if (rr !== r1 || cc !== c1) skip[rr + ',' + cc] = 1;
    });
    var h = '<table class="xl-grid"><colgroup><col class="xl-rn">';
    for (var c2 = 1; c2 <= maxC; c2++) { var w = ws.getColumn(c2).width; h += '<col style="width:' + Math.round(Math.max(18, Math.min(220, (w || 9) * 7.2))) + 'px">'; }
    h += '</colgroup><thead><tr><th class="xl-corner" aria-hidden="true"></th>';
    for (c2 = 1; c2 <= maxC; c2++) h += '<th scope="col" class="xl-col">' + colLetter(c2) + '</th>';
    h += '</tr></thead><tbody>';
    for (r = 1; r <= maxR; r++) {
      var rowH = ws.getRow(r).height;
      h += '<tr' + (rowH ? ' style="height:' + Math.round(rowH * 1.25) + 'px"' : '') + '><th scope="row" class="xl-rown">' + r + '</th>';
      for (c = 1; c <= maxC; c++) {
        if (skip[r + ',' + c]) continue;
        var cell = ws.getCell(r, c), st = [], cls = ['xl-v'];
        var f = cell.fill, fg = f && f.type === 'pattern' && f.pattern === 'solid' ? color(f.fgColor) : null;
        if (fg) st.push('background:' + fg);
        var fo = cell.font || {};
        if (fo.bold) st.push('font-weight:700');
        if (fo.italic) st.push('font-style:italic');
        var fc = color(fo.color); if (fc && fc !== '#FFFFFF' || (fc === '#FFFFFF' && fg)) st.push('color:' + fc);
        if (fo.size && fo.size >= 13) st.push('font-size:' + Math.round(fo.size * 1.05) + 'px');
        var al = cell.alignment || {};
        st.push('text-align:' + (al.horizontal === 'left' ? 'left' : al.horizontal === 'right' ? 'right' : al.horizontal === 'center' || al.horizontal === 'centerContinuous' ? 'center' : (typeof cell.value === 'number' || (cell.value && cell.value.formula) ? 'right' : 'left')));
        if (al.wrapText) cls.push('xl-wrap');
        var bd = cell.border || {};
        var bt = bw(bd.top), bb = bw(bd.bottom), bl = bw(bd.left), br = bw(bd.right);
        if (bt || bb || bl || br) st.push('border-style:solid;border-color:' + (bt ? '#000' : '#e1e1e1') + ' ' + (br ? '#000' : '#e1e1e1') + ' ' + (bb ? '#000' : '#e1e1e1') + ' ' + (bl ? '#000' : '#e1e1e1') + ';border-width:' + (bt || 1) + 'px ' + (br || 1) + 'px ' + (bb || 1) + 'px ' + (bl || 1) + 'px');
        var rot = al.textRotation && al.textRotation !== 0;
        var sp = span[r + ',' + c];
        var note = cell.note ? (typeof cell.note === 'string' ? cell.note : (cell.note.texts || []).map(function (t) { return t.text; }).join('')) : '';
        if (note) cls.push('xl-note');
        var val = esc(display(cell));
        h += '<td class="' + cls.join(' ') + '" style="' + st.join(';') + '"' + (sp ? (sp[0] > 1 ? ' rowspan="' + sp[0] + '"' : '') + (sp[1] > 1 ? ' colspan="' + sp[1] + '"' : '') : '') + (note ? ' title="' + esc(note) + '"' : '') + '>' +
          (rot ? '<span class="xl-rot">' + val + '</span>' : val) + (note ? '<span class="sr"> (' + esc(note) + ')</span>' : '') + '</td>';
      }
      h += '</tr>';
    }
    return h + '</tbody></table>';
  }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function colLetter(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  return { render: render };
})();
if (typeof module !== 'undefined') module.exports = Grid;
