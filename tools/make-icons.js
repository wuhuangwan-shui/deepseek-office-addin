'use strict';
/* 从 assets/icon.svg 重新生成插件图标 PNG（16/32/64/80）
   渲染：无头 Edge（4 倍超采样）→ System.Drawing 高质量缩小
   用法：node tools/make-icons.js   （或 npm run icons） */
var fs = require('fs');
var path = require('path');
var child = require('child_process');

var ROOT = path.resolve(__dirname, '..');
var SVG = path.join(ROOT, 'assets', 'icon.svg');
var SIZES = [16, 32, 64, 80];
var EDGE_CANDIDATES = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
];

function findEdge() {
  for (var i = 0; i < EDGE_CANDIDATES.length; i++) {
    if (fs.existsSync(EDGE_CANDIDATES[i])) return EDGE_CANDIDATES[i];
  }
  return null;
}

if (!fs.existsSync(SVG)) {
  console.error('找不到图标源文件: ' + SVG);
  process.exit(1);
}
var edge = findEdge();
if (!edge) {
  console.error('未找到 Edge 浏览器，无法渲染 SVG。请安装 Microsoft Edge 后重试。');
  process.exit(1);
}

var tmp = path.join(ROOT, '.icon-render');
fs.mkdirSync(tmp, { recursive: true });
var svgUri = 'file:///' + SVG.replace(/\\/g, '/');

function run(cmd) {
  return child.spawnSync(cmd[0], cmd.slice(1), { stdio: 'ignore', timeout: 60000 });
}

function downscale(srcPath, dstPath, size) {
  // 通过 PowerShell + System.Drawing 做高质量缩小（Edge 输出 4 倍图）
  var ps = [
    'Add-Type -AssemblyName System.Drawing;',
    '$src=[System.Drawing.Bitmap]::new("' + srcPath + '");',
    '$dst=[System.Drawing.Bitmap]::new(' + size + ',' + size + ',[System.Drawing.Imaging.PixelFormat]::Format32bppArgb);',
    '$g=[System.Drawing.Graphics]::FromImage($dst);',
    '$g.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic;',
    '$g.SmoothingMode=[System.Drawing.Drawing2D.SmoothingMode]::HighQuality;',
    '$g.Clear([System.Drawing.Color]::Transparent);',
    '$g.DrawImage($src,0,0,' + size + ',' + size + ');',
    '$g.Dispose();$src.Dispose();',
    '$dst.Save("' + dstPath + '",[System.Drawing.Imaging.ImageFormat]::Png);$dst.Dispose()',
  ].join('');
  var r = child.spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps], { timeout: 60000 });
  if (r.status !== 0) throw new Error('PowerShell 缩小失败（退出码 ' + r.status + '）');
}

SIZES.forEach(function (size) {
  var big = size * 4;
  var html = path.join(tmp, 's' + size + '.html');
  fs.writeFileSync(html, '<html><body style="margin:0;padding:0;background:transparent">' +
    '<img src="' + svgUri + '" width="' + big + '" height="' + big + '"></body></html>', 'utf8');
  var raw = path.join(tmp, 'raw-' + size + '.png');
  run([edge, '--headless=new', '--no-sandbox', '--disable-gpu', '--no-first-run',
    '--force-device-scale-factor=1', '--default-background-color=00000000',
    '--window-size=' + big + ',' + big,
    '--screenshot=' + raw, 'file:///' + html.replace(/\\/g, '/')]);
  if (!fs.existsSync(raw)) {
    console.error('渲染失败: ' + size + 'px（Edge 无输出）');
    process.exit(1);
  }
  var out = path.join(ROOT, 'assets', 'icon-' + size + '.png');
  downscale(raw, out, size);
  console.log('已生成 assets/icon-' + size + '.png');
});

fs.rmSync(tmp, { recursive: true, force: true });
console.log('完成。图标源文件: assets/icon.svg');
