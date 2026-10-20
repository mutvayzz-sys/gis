// Builds public/index.html from the files in src/. Run: npm run build
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const src = (f) => fs.readFileSync(path.join(root, 'src', f), 'utf8');

let page = src('template.html');
const fill = {
  EDITOR: '', N_ALL: '0', N_GRADES: '0', N_TOTALS: '0', DIST_NOW: '', TABLE_NOW: '', FIXES: '', COMPARE: '',
  CORE: src('core.js').replace("if (typeof module !== 'undefined') module.exports = Core;", ''),
  VIEW: src('view.js').replace("if (typeof module !== 'undefined') module.exports = View;", ''),
  APP: src('xlwriter.js') + '\n' + src('app.js')
};
page = page.replace(/\{\{([A-Z_]+)\}\}/g, (m, k) => {
  if (fill[k] == null) throw new Error('No value for placeholder ' + k);
  return fill[k];
});
page = page.replace('<a class="skip"', '<script>document.body.classList.add("is-empty")</script>\n<a class="skip"');

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
