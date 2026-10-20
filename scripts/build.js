// Builds public/index.html from src/. Run: npm run build
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const rd = (f) => fs.readFileSync(path.join(root, 'src', f), 'utf8');
const strip = (s, name) => s.replace("if (typeof module !== 'undefined') module.exports = " + name + ";", '');

const fill = {
  PROMPT: JSON.stringify(rd('prompt.txt')),
  ENGINE: strip(rd('engine.js'), 'Engine'),
  BOOK: strip(rd('book.js'), 'Book'),
  GRID: strip(rd('grid.js'), 'Grid'),
  APP: rd('app.js').replace('/*ASSISTANT*/', () => rd('assistant.js')).replace('/*REPORT*/', () => rd('report.js'))
};
const page = rd('page.html').replace(/\{\{([A-Z]+)\}\}/g, (m, k) => {
  if (fill[k] == null) throw new Error('No value for placeholder ' + k);
  return fill[k];
});
const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="robots" content="noindex, nofollow">
<style>:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;font-size:14px}img{max-width:100%}[hidden]{display:none!important}</style>
</head>
<body>
${page}
</body>
</html>
`;
fs.mkdirSync(path.join(root, 'public'), { recursive: true });
fs.writeFileSync(path.join(root, 'public', 'index.html'), html);
console.log('Built public/index.html (' + Math.round(html.length / 1024) + ' KB)');
