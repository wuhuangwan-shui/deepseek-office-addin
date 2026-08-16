/**
 * DeepSeek 办公助手 - 本地服务
 * 1) 托管插件静态页面（taskpane.html 等），HTTPS（命令型加载项强制要求）
 * 2) 转发 DeepSeek API 请求（避免浏览器跨域，支持流式输出）
 * 运行：node server.js（端口默认 3000，可用环境变量 PORT 覆盖）
 * 首次运行会自动调用 setup-certs.ps1 生成本地证书并加入当前用户信任根。
 */
'use strict';

var https = require('https');
var fs = require('fs');
var path = require('path');
var child = require('child_process');

var PORT = Number(process.env.PORT || 3000);
var ROOT = __dirname;
var PFX_PATH = path.join(ROOT, 'cert', 'localhost.pfx');
var PFX_PASS = 'deepseek-dev';

function ensureHttpsCert() {
  if (process.env.OFFICE_ADDIN_SKIP_CERT_CHECK) return;
  // 每次启动都跑 setup-certs.ps1：缺证书则生成，已存在则只做信任检查（幂等）
  console.log('检查本地 HTTPS 证书...');
  var r = child.spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(ROOT, 'setup-certs.ps1')], { stdio: 'inherit', timeout: 60000 });
  if (!fs.existsSync(PFX_PATH)) {
    console.error('HTTPS 证书文件缺失（退出码 ' + r.status + '）。请手动用 PowerShell 运行 setup-certs.ps1 查看原因。');
    process.exit(1);
  }
  if (r.status !== 0) {
    console.warn('警告: 证书信任检查未通过。Word/Excel 加载插件时可能报错，请重新运行 setup-certs.ps1 或安装到Office.cmd 完成信任。');
  }
}
var MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8',
  '.png': 'image/png',
  '.md': 'text/plain; charset=utf-8',
  '.ico': 'image/x-icon',
};

function send(res, status, obj) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(obj));
}

function serveStatic(req, res, urlPath) {
  if (urlPath === '/' || urlPath === '/index.html') urlPath = '/taskpane.html';
  var file = path.normalize(path.join(ROOT, urlPath));
  if (file.slice(0, ROOT.length) !== ROOT) return send(res, 403, { error: 'forbidden' });
  fs.readFile(file, function (err, data) {
    if (err) return send(res, 404, { error: 'not found: ' + urlPath });
    var ext = path.extname(file).toLowerCase();
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
    res.end(data);
  });
}

function readBody(req, limit) {
  return new Promise(function (resolve, reject) {
    var size = 0;
    var chunks = [];
    req.on('data', function (c) {
      size += c.length;
      if (size > limit) {
        reject(new Error('body too large'));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on('end', function () { resolve(Buffer.concat(chunks).toString('utf8')); });
    req.on('error', reject);
  });
}

function handleChat(req, res) {
  var started = false; // 是否已开始向客户端回包
  readBody(req, 20 * 1024 * 1024).then(function (raw) {
    var payload;
    try { payload = JSON.parse(raw); } catch (e) { return send(res, 400, { error: { message: '无效的请求体' } }); }
    var apiKey = req.headers['x-api-key'] || process.env.DEEPSEEK_API_KEY || '';
    if (!apiKey) {
      return send(res, 400, { error: { message: '未配置 API Key：请在任务窗格「⚙️ 设置」中填写，或设置环境变量 DEEPSEEK_API_KEY 后重启本服务' } });
    }
    var body = {
      model: payload.model || 'deepseek-v4-flash',
      messages: Array.isArray(payload.messages) ? payload.messages : [],
      temperature: (payload.temperature === undefined || payload.temperature === null) ? 0.7 : payload.temperature,
      stream: !!payload.stream,
    };
    if (payload.max_tokens) body.max_tokens = Math.min(Number(payload.max_tokens) || 0, 8192);
    var rawBody = JSON.stringify(body);
    var up = https.request({
      hostname: 'api.deepseek.com',
      path: '/chat/completions',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey,
        'Content-Length': Buffer.byteLength(rawBody),
      },
    }, function (upr) {
      started = true;
      res.writeHead(upr.statusCode || 502, {
        'Content-Type': upr.headers['content-type'] || 'application/json; charset=utf-8',
        'Cache-Control': 'no-cache',
      });
      upr.pipe(res);
    });
    // 上游超时保护：DeepSeek 长时间无响应时中止，避免任务窗格一直等待
    up.setTimeout(120000, function () {
      up.destroy(new Error('DeepSeek 服务器响应超时'));
    });
    up.on('error', function (e) {
      if (started) { res.destroy(); return; } // 已开始回包则直接断开流
      send(res, 502, { error: { message: '无法连接 DeepSeek 服务器：' + e.message } });
    });
    up.end(rawBody);
  }).catch(function (e) {
    send(res, 400, { error: { message: '读取请求失败：' + e.message } });
  });
}

ensureHttpsCert();
if (process.env.OFFICE_ADDIN_SKIP_CERT_CHECK) console.log('（OFFICE_ADDIN_SKIP_CERT_CHECK=1：已跳过启动时的证书信任检查）');

var server = https.createServer({ pfx: fs.readFileSync(PFX_PATH), passphrase: PFX_PASS }, function (req, res) {
  var url = new URL(req.url, 'http://localhost');
  if (req.method === 'POST' && url.pathname === '/api/chat') return handleChat(req, res);
  if (req.method === 'GET' && url.pathname === '/api/config') {
    return send(res, 200, { serverKeySet: !!process.env.DEEPSEEK_API_KEY, port: PORT });
  }
  if (req.method === 'GET' || req.method === 'HEAD') return serveStatic(req, res, url.pathname);
  send(res, 404, { error: 'not found' });
});

// 端口占用等启动错误：给出友好提示而不是崩溃堆栈
server.on('error', function (e) {
  if (e && e.code === 'EADDRINUSE') {
    console.error('启动失败：端口 ' + PORT + ' 已被占用。');
    console.error('可能原因：另一个 start.cmd 已在运行，或该端口被其他程序使用。');
    console.error('处理：关闭多余的 start.cmd 窗口后重试；如需换端口，设置环境变量 PORT（如 3001）并同步修改 manifest.xml 中的地址。');
  } else if (e && e.code === 'EACCES') {
    console.error('启动失败：没有权限监听端口 ' + PORT + '。请换一个高位端口（如 3000 以上）。');
  } else {
    console.error('启动失败：' + (e && e.message ? e.message : e));
  }
  process.exit(1);
});

server.listen(PORT, function () {
  console.log('DeepSeek 办公助手服务已启动: https://localhost:' + PORT);
  console.log('插件清单: https://localhost:' + PORT + '/manifest.xml');
  if (PORT !== 3000) {
    console.log('注意: 端口不是 3000，请同步修改 manifest.xml 中的 https://localhost:3000 为 https://localhost:' + PORT);
  }
  console.log(process.env.DEEPSEEK_API_KEY ? '已检测到环境变量 DEEPSEEK_API_KEY（任务窗格可免填 Key）' : '未检测到 DEEPSEEK_API_KEY，请在任务窗格 ⚙️ 设置中填写');
});
