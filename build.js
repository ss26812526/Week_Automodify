// 把 src/index.html 及其相依合成單一自足的 HTML 發布檔
// 用法：node build.js
var fs = require('fs');
var path = require('path');

var root = __dirname;
var srcDir = path.join(root, 'src');
var outDir = path.join(root, 'dist');
var outFile = path.join(outDir, '檢查表星期填寫工具.html');

function readRelative(relative) {
  var p = path.resolve(srcDir, relative);
  if (!fs.existsSync(p)) throw new Error('找不到檔案：' + p);
  return fs.readFileSync(p, 'utf8');
}

var html = fs.readFileSync(path.join(srcDir, 'index.html'), 'utf8');

html = html.replace(/<link\s+rel="stylesheet"\s+href="([^"]+)"\s*>/g, function (_, href) {
  console.log('內嵌樣式 ' + href);
  return '<style>\n' + readRelative(href) + '\n</style>';
});

html = html.replace(/<script\s+src="([^"]+)"><\/script>/g, function (_, src) {
  console.log('內嵌腳本 ' + src);
  var code = readRelative(src);
  return '<script>\n' + code.replace(/<\/script>/gi, '<\\/script>') + '\n</script>';
});

var marker = '<!-- 由 build.js 合成，請勿直接編輯。原始碼在 src/ 目錄。 -->\n';
html = html.replace('</head>', marker + '</head>');

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outFile, html, 'utf8');

var kb = Math.round(fs.statSync(outFile).size / 1024);
console.log('');
console.log('已產生 ' + outFile + '（' + kb + ' KB）');
if (html.indexOf('<script src=') !== -1 || html.indexOf('<link rel="stylesheet"') !== -1) {
  throw new Error('仍有未內嵌的外部檔案，發布檔不是自足的');
}
console.log('確認：檔案中已無任何外部相依');
