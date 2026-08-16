/* global Office, Word, Excel */
'use strict';

var NL = String.fromCharCode(10);
var CR = String.fromCharCode(13);
var BT = String.fromCharCode(96);
var LS_KEY = 'dsOfficeSettings';
var MAX_CELLS = 12000;

var state = {
  host: null,
  model: 'deepseek-v4-flash',
  temperature: 0.7,
  currentAction: null,
  busy: false,
};

function $(sel) { return document.querySelector(sel); }

/* ---------------- 设置 ---------------- */
function loadSettings() {
  try {
    var s = JSON.parse(localStorage.getItem(LS_KEY) || '{}');
    // 旧模型名迁移（2026-07 官方停用 deepseek-chat / deepseek-reasoner）
    if (s.model === 'deepseek-chat') s.model = 'deepseek-v4-flash';
    if (s.model === 'deepseek-reasoner') s.model = 'deepseek-v4-pro';
    if (s.model) state.model = s.model;
    if (s.temperature !== null && s.temperature !== undefined) state.temperature = Number(s.temperature);
    if (s.apiKey) $('#apiKey').value = s.apiKey;
  } catch (e) {}
}

function saveSettings() {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify({
      model: state.model,
      temperature: state.temperature,
      apiKey: $('#apiKey').value.trim(),
    }));
  } catch (e) {}
}

/* ---------------- 小工具 ---------------- */
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function toast(msg) {
  var t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast._timer);
  toast._timer = setTimeout(function () { t.classList.remove('show'); }, 2800);
}

function chatScroll() {
  var c = $('#chat');
  c.scrollTop = c.scrollHeight;
}

function setBusy(b) {
  state.busy = b;
  $('#btnSend').disabled = b;
  var cards = document.querySelectorAll('.action-card');
  for (var i = 0; i < cards.length; i++) cards[i].disabled = b;
}

function cellText(c) {
  if (c === null || c === undefined) return '';
  return String(c).split('|').join(' ').split(NL).join(' ').split(CR).join(' ');
}

function rowsToMarkdown(rows) {
  if (!rows || !rows.length) return '(空选区)';
  var lines = [];
  var header = rows[0].map(function (c) { return cellText(c); });
  lines.push('| ' + header.join(' | ') + ' |');
  lines.push('|' + header.map(function () { return '---'; }).join('|') + '|');
  for (var i = 1; i < rows.length; i++) {
    lines.push('| ' + rows[i].map(function (c) { return cellText(c); }).join(' | ') + ' |');
  }
  var text = lines.join(NL);
  if (text.length > MAX_CELLS) text = text.slice(0, MAX_CELLS) + NL + '…（内容过长已截断）';
  return text;
}

/* ---------------- 聊天 UI ---------------- */
function addMessage(role, text) {
  var empty = document.querySelector('#chat .empty');
  if (empty && empty.parentNode) empty.parentNode.removeChild(empty);
  var chat = $('#chat');
  var msg = document.createElement('div');
  msg.className = 'msg ' + role;
  var body = document.createElement('div');
  body.className = 'body';
  body.textContent = text || '';
  msg.appendChild(body);
  chat.appendChild(msg);
  chatScroll();
  return { msg: msg, body: body };
}

function addApplyButton(msg, text, writable, label) {
  var bar = document.createElement('div');
  bar.className = 'msg-tools';
  var btn = document.createElement('button');
  btn.className = 'mini';
  btn.textContent = label;
  btn.onclick = function () {
    btn.disabled = true;
    btn.textContent = '正在写入…';
    applyToDocument(writable, text).then(function () {
      btn.textContent = '已完成';
    }).catch(function (e) {
      btn.textContent = '重试';
      toast('写入失败：' + (e && e.message ? e.message : e));
      btn.disabled = false;
    });
  };
  bar.appendChild(btn);
  msg.appendChild(bar);
  chatScroll();
}

/* ---------------- 调用 DeepSeek（经本地服务转发） ---------------- */
function runChat(messages, onDelta, onReasoning, maxTokens) {
  var apiKey = $('#apiKey').value.trim();
  return fetch('/api/config')
    .then(function (r) { return r.json(); })
    .catch(function () { return {}; })
    .then(function (cfg) {
      if (!apiKey && !(cfg && cfg.serverKeySet)) {
        throw new Error('请先在「设置」中填写 DeepSeek API Key（或为服务设置环境变量 DEEPSEEK_API_KEY）');
      }
      var payload = {
        model: state.model,
        messages: messages,
        temperature: state.temperature,
        stream: true,
        max_tokens: maxTokens || 8000,
      };
      return fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Api-Key': apiKey },
        body: JSON.stringify(payload),
      }).then(function (resp) {
        if (!resp.ok) {
          return resp.text().then(function (t) {
            var msg = '请求失败 (HTTP ' + resp.status + ')';
            try {
              var j = JSON.parse(t);
              if (j && j.error && j.error.message) msg += '：' + j.error.message;
            } catch (e) {}
            throw new Error(msg);
          });
        }
        return resp;
      });
    })
    .then(function (resp) {
      var reader = resp.body.getReader();
      var decoder = new TextDecoder();
      var buf = '';
      function pump() {
        return reader.read().then(function (r) {
          if (r.done) return;
          buf += decoder.decode(r.value, { stream: true });
          var idx;
          while ((idx = buf.indexOf(NL)) >= 0) {
            var line = buf.slice(0, idx).trim();
            buf = buf.slice(idx + 1);
            if (line.indexOf('data:') !== 0) continue;
            var data = line.slice(5).trim();
            if (data === '[DONE]') continue;
            try {
              var j = JSON.parse(data);
              var d = j.choices && j.choices[0] && j.choices[0].delta;
              if (!d) continue;
              if (d.reasoning_content && onReasoning) onReasoning(d.reasoning_content);
              if (d.content && onDelta) onDelta(d.content);
            } catch (e) {}
          }
          return pump();
        });
      }
      return pump();
    });
}

function streamAssistant(cfg) {
  var content = '';
  var reasoning = '';
  var reasoningEl = null;
  var msgObj = addMessage('assistant', '');
  var msg = msgObj.msg;
  var body = msgObj.body;
  var messages = [
    { role: 'system', content: cfg.system },
    { role: 'user', content: cfg.user },
  ];
  var retried = false;

  function runOnce(maxTokens, showReasoning) {
    return runChat(messages, function (delta) {
      content += delta;
      body.textContent = content;
      chatScroll();
    }, showReasoning ? function (delta) {
      reasoning += delta;
      if (!reasoningEl) {
        var details = document.createElement('details');
        details.className = 'thinking';
        var sum = document.createElement('summary');
        sum.textContent = '思考过程';
        var pre = document.createElement('div');
        details.appendChild(sum);
        details.appendChild(pre);
        msg.insertBefore(details, body);
        reasoningEl = pre;
      }
      reasoningEl.textContent = reasoning;
      chatScroll();
    } : null, maxTokens);
  }

  function finish() {
    if (!content && reasoning) {
      body.textContent = '（模型只输出了思考过程，自动重试后仍无结果。任务可能过于复杂，可换用 deepseek-v4-flash 或重试）';
    }
    if (cfg.writable && content.trim()) {
      addApplyButton(msg, content.trim(), cfg.writable, cfg.applyLabel || '写入文档');
    }
    chatScroll();
  }

  return runOnce(cfg.maxTokens || 8000, true).then(function () {
    if (!content && reasoning && !retried) {
      // 思考耗尽了输出预算：清掉思考展示，用最大预算自动重试一次（不再展示思考）
      retried = true;
      reasoning = '';
      if (reasoningEl && reasoningEl.parentNode) reasoningEl.parentNode.removeChild(reasoningEl);
      reasoningEl = null;
      body.textContent = '';
      return runOnce(8000, false).then(finish);
    }
    finish();
  }).catch(function (e) {
    body.textContent = '出错了：' + (e && e.message ? e.message : e);
    chatScroll();
  });
}

/* ---------------- 能力定义 ---------------- */
var ACTIONS = {
  Word: [
    {
      id: 'formatPaper', label: '论文排版', desc: '一键应用标准中文论文格式', icon: 'E8E4', run: formatPaperDocument,
    },
    {
      id: 'write', label: '写作', desc: '输入要求，内容插入光标处', icon: 'E70F', requiresPrompt: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是一名专业的中文办公写作助手。根据用户的要求直接输出最终成稿：不要任何解释、前后缀或客套话；内容可直接插入 Word 文档，需要分段或小标题时按合适格式输出。',
          user: user,
        };
      },
    },
    {
      id: 'polish', label: '润色改写', desc: '优化选中文字的措辞', icon: 'E70B', needsSelection: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是文字润色专家。请改写用户提供的文本：保持原意，让表达更准确、专业、流畅。只输出改写后的文本，不要任何解释。',
          user: (user ? '润色要求：' + user + NL + NL : '') + '待润色文本：' + NL + ctx.selectedText,
        };
      },
    },
    {
      id: 'rewrite', label: '换种说法', desc: '用不同表达重写选中文字', icon: 'E72C', needsSelection: true, writable: 'replace',
      build: function (ctx, user) {
        return {
          system: '你是写作助手。请用与原文明显不同的表达方式重写用户提供的文本，保持原意不变。只输出重写后的文本。',
          user: (user ? '风格要求：' + user + NL + NL : '') + '原文：' + NL + ctx.selectedText,
        };
      },
    },
    {
      id: 'expand', label: '扩写', desc: '把选中文字写得更充实', icon: 'E8A3', needsSelection: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是写作助手。请扩写用户提供的文本：保持原意和语气，补充细节、论证或过渡，使内容更充实完整。只输出扩写后的全文。',
          user: (user ? '扩写要求：' + user + NL + NL : '') + '原文：' + NL + ctx.selectedText,
        };
      },
    },
    {
      id: 'shorten', label: '缩写', desc: '压缩选中文字，保留要点', icon: 'E71F', needsSelection: true, writable: 'replace',
      build: function (ctx, user) {
        return {
          system: '你是编辑助手。请缩写用户提供的文本：保留核心信息和逻辑，删除冗余，压缩到原文的 30%-50%。只输出缩写后的文本。',
          user: (user ? '要求：' + user + NL + NL : '') + '原文：' + NL + ctx.selectedText,
        };
      },
    },
    {
      id: 'summary', label: '总结', desc: '总结选中文字的要点', icon: 'E8FD', needsSelection: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是总结助手。请用中文总结用户提供的文本要点：先给一句总括，再列 3-6 个要点。简洁、结构化。只输出总结内容。',
          user: '待总结文本：' + NL + ctx.selectedText,
        };
      },
    },
    {
      id: 'docSummary', label: '全文总结', desc: '总结整篇文档', icon: 'E8A5', useWholeDoc: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是总结助手。请总结用户提供的整篇文档：主题、结构、关键结论。用中文分点输出。',
          user: '文档全文（可能被截断）：' + NL + (ctx.docText || '(文档为空)'),
        };
      },
    },
    {
      id: 'translate', label: '翻译', desc: '翻译选中文字（中英互译）', icon: 'E909', needsSelection: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是翻译专家。把用户提供的文本翻译成中文；如果原文已经是中文，则翻译成英文。只输出译文，不要任何解释。',
          user: ctx.selectedText,
        };
      },
    },
    {
      id: 'outline', label: '生成大纲', desc: '按主题生成文档大纲', icon: 'E7C3', requiresPrompt: true, writable: 'insert',
      build: function (ctx, user) {
        return {
          system: '你是文档策划专家。根据用户主题生成一份结构清晰的文档大纲：使用多级标题和要点列表。只输出大纲。',
          user: user,
        };
      },
    },
    {
      id: 'proofread', label: '校对纠错', desc: '检查选中文字的错别字和语病', icon: 'E721', needsSelection: true,
      build: function (ctx, user) {
        return {
          system: '你是校对专家。检查用户文本中的错别字、语法错误、标点误用、用词不当。先逐条列出「原文 → 修改建议」，最后给出修改后的全文。',
          user: ctx.selectedText,
        };
      },
    },
  ],
  Excel: [
    {
      id: 'analyze', label: '分析数据', desc: '分析选中区域，找规律查问题', icon: 'E80F', needsData: true,
      build: function (ctx, user) {
        return {
          system: '你是数据分析师。用户提供了一段 Excel 选区数据（Markdown 表格，首行为表头）。请用中文输出：1) 数据概况 2) 关键数字与规律 3) 发现的问题或异常 4) 可操作的建议。分点简洁输出。',
          user: '选区：' + ctx.address + NL + (user ? '分析要求：' + user + NL + NL : '') + ctx.markdown,
        };
      },
    },
    {
      id: 'formula', label: '生成公式', desc: '一句话描述需求，公式写入选中单元格', icon: 'E8EF', requiresPrompt: true, writable: 'formula',
      build: function (ctx, user) {
        return {
          system: '你是 Excel 公式专家。根据用户描述生成一个 Excel 公式。要求：只输出公式本身，以 = 开头；不要代码块标记、不要解释。函数名使用英文（SUM、IF、IFS、VLOOKUP、XLOOKUP、INDEX、MATCH、COUNTIF、TEXTJOIN、LEFT、RIGHT、MID 等）。',
          user: user + NL + '当前选中单元格：' + ctx.address + NL + (ctx.markdown !== '(空选区)' ? '相关数据（首行为表头）：' + NL + ctx.markdown : ''),
        };
      },
    },
    {
      id: 'fill', label: '智能填充', desc: '按选中数据规律续写，从选中单元格写入', icon: 'E896', needsData: true, writable: 'values', maxTokens: 8000,
      build: function (ctx, user) {
        return {
          system: '你是 Excel 助手。用户给出已有数据（Markdown 表格，首行为表头），要求生成数据。请只输出 JSON 数组（每行一个数组，数字不带引号，文本带引号），不要输出任何其他文字、注释或代码块标记。',
          user: (user || '请根据已有数据的规律，继续生成 10 行新数据。') + NL + '已有数据：' + NL + ctx.markdown,
        };
      },
    },
    {
      id: 'clean', label: '数据清洗', desc: '分类、规范、去重等，覆盖写回选中区域', icon: 'E75C', needsData: true, writable: 'values', maxTokens: 8000,
      build: function (ctx, user) {
        return {
          system: '你是 Excel 数据清洗助手。用户给出表格（首行为表头）和清洗要求。请只输出清洗后的完整数据：JSON 数组，每行一个数组，与原行一一对应（行数一致），数字不带引号，文本带引号。不要输出任何解释或代码块标记。',
          user: (user || '请规范日期格式、去除多余空格、统一文本大小写。') + NL + '数据：' + NL + ctx.markdown,
        };
      },
    },
    {
      id: 'translateCells', label: '翻译单元格', desc: '把选中区域文本翻译成中文并写回', icon: 'E909', needsData: true, writable: 'values', maxTokens: 8000,
      build: function (ctx, user) {
        return {
          system: '你是翻译助手。用户给出 Excel 表格（首行为表头）。请把每个单元格的文本翻译成中文（数字、公式、已是中文的内容保持原样）。只输出 JSON 数组：每行一个数组，与原表行数、列数完全一致。不要输出任何解释。',
          user: '数据：' + NL + ctx.markdown,
        };
      },
    },
    {
      id: 'extract', label: '提取信息', desc: '从每行提取邮箱/电话等，写入选中区域', icon: 'E779', needsData: true, writable: 'values', maxTokens: 8000,
      build: function (ctx, user) {
        return {
          system: '你是 Excel 助手。用户给出表格和要求。请从每行数据中提取要求的信息，输出 JSON 数组：每行一个数组，每个元素是该行的提取结果（提取不到则用空字符串）。只输出 JSON。',
          user: (user || '请从每行文本中提取邮箱地址和手机号。') + NL + '数据：' + NL + ctx.markdown,
        };
      },
    },
  ],
};

function findAction(id) {
  var list = ACTIONS[state.host] || [];
  for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
  return null;
}

/* ---------------- 读取文档上下文 ---------------- */
// Excel 日期序列 → 'YYYY-MM-DD'（1899-12-30 基准）
function excelSerialToDate(n) {
  var ms = Date.UTC(1899, 11, 30) + Math.round(n * 86400000);
  var d = new Date(ms);
  function p(x) { return (x < 10 ? '0' : '') + x; }
  return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate());
}

// 日期格式单元格的序列值转可读日期（发给 AI 前）
function fmtCell(v, fmt) {
  if (typeof v !== 'number' || !fmt) return v;
  var f = String(fmt);
  if (/[hHsS]/.test(f)) return v; // 含时间成分 → 保持原值
  if (/[yY]/.test(f) && /[mM]/.test(f) && /[dD]/.test(f) && v > 10000 && v < 80000) {
    return excelSerialToDate(v);
  }
  return v;
}

function collectWordContext(action) {
  return Word.run(function (context) {
    var info = { host: 'Word' };
    var sel = null, body = null;
    if (action.needsSelection) { sel = context.document.getSelection(); sel.load('text'); }
    if (action.useWholeDoc) { body = context.document.body; body.load('text'); }
    return context.sync().then(function () {
      if (sel) {
        info.selectedText = (sel.text || '').trim();
        info.emptySelection = !info.selectedText;
      }
      if (body) info.docText = (body.text || '').slice(0, 20000);
      return info;
    });
  });
}

function collectExcelContext(action) {
  return Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(['values', 'numberFormat', 'rowCount', 'columnCount', 'address']);
    return context.sync().then(function () {
      var info = { host: 'Excel', address: range.address, rowCount: range.rowCount, columnCount: range.columnCount };
      var fmts = range.numberFormat || [];
      var rows = range.values.slice(0, 200).map(function (r, ri) {
        return r.slice(0, 50).map(function (v, ci) { return fmtCell(v, fmts[ri] && fmts[ri][ci]); });
      });
      info.table = rows;
      info.markdown = rowsToMarkdown(rows);
      if (range.rowCount > 200 || range.columnCount > 50) info.markdown += NL + '…（选区过大，仅取了前 200 行 × 50 列）';
      return info;
    });
  });
}

function collectContext(action) {
  if (state.host === 'Word') return collectWordContext(action);
  return collectExcelContext(action);
}

/* ---------------- 结构感知写入（Markdown → Word HTML） ---------------- */
// 输出内容是否包含需要结构保留的 Markdown 特征
function hasMarkdownStructure(text) {
  return /^#{1,6}\s/m.test(text)        // 标题
    || /^[-+*]\s/m.test(text)           // 无序列表
    || /^\d+[.)]\s/m.test(text)         // 有序列表
    || /\*\*[^*]+\*\*/.test(text)       // 加粗
    || /^\|.*\|\s*$/m.test(text)        // 表格行
    || /^```/m.test(text)               // 代码块
    || /^>\s?/m.test(text);             // 引用
}

function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// 行内标记 → HTML（加粗/斜体/行内代码/链接）
function inlineHtml(s) {
  var t = escHtml(s);
  t = t.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  t = t.replace(/__([^_]+)__/g, '<strong>$1</strong>');
  t = t.replace(/\*([^*\n]+)\*/g, '<em>$1</em>');
  t = t.replace(/\`([^`\n]+)\`/g, '<code>$1</code>');
  t = t.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, '<a href="$2">$1</a>');
  return t;
}

// 简易 Markdown → Word 可导入的 HTML（标题/列表/表格/代码块/引用/段落）
function markdownToHtml(text) {
  var lines = String(text).split(NL);
  var html = [];
  var list = null; // 'ul' | 'ol' | null
  function closeList() {
    if (list) { html.push('</' + list + '>'); list = null; }
  }
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) { closeList(); continue; }
    var m;
    if ((m = line.match(/^(#{1,6})\s+(.*)$/))) {
      closeList();
      var lvl = m[1].length;
      html.push('<h' + lvl + '>' + inlineHtml(m[2]) + '</h' + lvl + '>');
      continue;
    }
    if (/^```/.test(line)) {
      closeList();
      var code = [];
      i++;
      while (i < lines.length && !/^```/.test(lines[i].trim())) { code.push(escHtml(lines[i])); i++; }
      html.push('<pre>' + code.join('<br/>') + '</pre>');
      continue;
    }
    if ((m = line.match(/^\s*[-+*]\s+(.*)$/))) {
      if (list !== 'ul') { closeList(); html.push('<ul>'); list = 'ul'; }
      html.push('<li>' + inlineHtml(m[1]) + '</li>');
      continue;
    }
    if ((m = line.match(/^\s*(\d+)[.)]\s+(.*)$/))) {
      if (list !== 'ol') { closeList(); html.push('<ol>'); list = 'ol'; }
      html.push('<li>' + inlineHtml(m[2]) + '</li>');
      continue;
    }
    if (/^\|.*\|\s*$/.test(line)) {
      closeList();
      var rows = [];
      while (i < lines.length && /^\|.*\|\s*$/.test(lines[i].trim())) { rows.push(lines[i].trim()); i++; }
      i--;
      var tbl = ['<table>'];
      rows.forEach(function (r, ri) {
        var cells = r.replace(/^\||\|$/g, '').split('|').map(function (c) { return c.trim(); });
        if (cells.every(function (c) { return /^:?-{2,}:?$/.test(c); })) return; // 分隔行
        var tag = ri === 0 ? 'th' : 'td';
        tbl.push('<tr>' + cells.map(function (c) { return '<' + tag + '>' + inlineHtml(c) + '</' + tag + '>'; }).join('') + '</tr>');
      });
      tbl.push('</table>');
      html.push(tbl.join(''));
      continue;
    }
    if (/^>\s?/.test(line)) {
      closeList();
      html.push('<blockquote>' + inlineHtml(line.replace(/^>\s?/, '')) + '</blockquote>');
      continue;
    }
    closeList();
    html.push('<p>' + inlineHtml(line) + '</p>');
  }
  closeList();
  return html.join('');
}

/* ---------------- 写回文档 ---------------- */
function parseJsonArray(text) {
  var t = String(text).trim();
  var fence = BT + BT + BT;
  var fidx = t.indexOf(fence);
  if (fidx >= 0) {
    // 取代码围栏内部内容（容忍围栏前后有说明文字）
    var after = t.slice(fidx + 3);
    var eidx = after.indexOf(fence);
    t = (eidx >= 0) ? after.slice(0, eidx) : after;
  }
  t = t.trim();
  var start = t.indexOf('[');
  var end = t.lastIndexOf(']');
  if (start < 0 || end <= start) return null;
  try {
    var arr = JSON.parse(t.slice(start, end + 1));
    return Array.isArray(arr) ? arr : null;
  } catch (e) {
    return null;
  }
}

function cleanFormula(text) {
  var t = String(text).trim();
  var fence = BT + BT + BT;
  var fidx = t.indexOf(fence);
  if (fidx >= 0) {
    var after = t.slice(fidx + 3);
    var eidx = after.indexOf(fence);
    t = (eidx >= 0) ? after.slice(0, eidx) : after;
  }
  t = t.replace(/\`/g, ''); // 清理残留的行内反引号
  var lines = t.split(NL);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line) { t = line; break; } // 跳过空行，取第一个非空行
  }
  t = t.trim();
  if (t && t.charAt(0) !== '=') t = '=' + t;
  return t;
}

function applyToDocument(writable, text) {
  // 注意：writable 是写入类型（insert/replace/formula/values），不是动作 ID
  if (state.host === 'Word') {
    // 输出含 Markdown 结构时用 HTML 写入以保留结构；纯文本用 insertText 保留原文段落格式
    var html = hasMarkdownStructure(text) ? markdownToHtml(text) : null;
    return Word.run(function (context) {
      var sel = context.document.getSelection();
      if (writable === 'replace') {
        if (html) sel.insertHtml(html, Word.InsertLocation.replace);
        else sel.insertText(text, Word.InsertLocation.replace);
      } else {
        if (html) sel.insertHtml(html, Word.InsertLocation.end);
        else sel.insertText(NL + text, Word.InsertLocation.end);
      }
      return context.sync();
    });
  }
  if (writable === 'formula') {
    var f = cleanFormula(text);
    return Excel.run(function (context) {
      var range = context.workbook.getSelectedRange();
      range.formulas = [[f]];
      return context.sync();
    });
  }
  var rows = parseJsonArray(text);
  if (!rows || !rows.length) throw new Error('AI 返回的内容不是表格数据，无法写入');
  return Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var width = 1;
    rows.forEach(function (r) {
      var w = Array.isArray(r) ? r.length : 1;
      if (w > width) width = w;
    });
    var data = rows.map(function (r) {
      var arr = Array.isArray(r) ? r : [r];
      var row = [];
      for (var i = 0; i < width; i++) {
        var v = arr[i];
        row.push((v === null || v === undefined || v === '') ? '' : v);
      }
      return row;
    });
    var target = range.getCell(0, 0).getResizedRange(data.length - 1, width - 1);
    target.values = data;
    return context.sync();
  });
}

/* ---------------- 论文排版（纯 Word API，不使用 AI） ---------------- */
// 标准中文论文格式规范（通用高校/期刊规范，可在此调整）
var PAPER_SPECS = {
  title: { size: 22, bold: true, east: '黑体', latin: 'Times New Roman', center: true, spaceAfter: 12, outline: 'OutlineLevel1' },
  author: { size: 12, bold: false, east: '宋体', latin: 'Times New Roman', center: true, spaceAfter: 6 },
  h1: { size: 16, bold: true, east: '黑体', latin: 'Times New Roman', center: true, spaceBefore: 12, spaceAfter: 6, outline: 'OutlineLevel1' },
  h2: { size: 14, bold: true, east: '黑体', latin: 'Times New Roman', center: false, spaceBefore: 6, spaceAfter: 3, outline: 'OutlineLevel2' },
  h3: { size: 12, bold: true, east: '黑体', latin: 'Times New Roman', center: false, spaceBefore: 3, outline: 'OutlineLevel3' },
  abstract: { size: 12, east: '宋体', latin: 'Times New Roman', justify: true, indent: 24 },
  keywords: { size: 12, east: '宋体', latin: 'Times New Roman', justify: true, indent: 24 },
  body: { size: 12, east: '宋体', latin: 'Times New Roman', justify: true, indent: 24 },
  ref: { size: 10.5, east: '宋体', latin: 'Times New Roman', hanging: 21 },
};

// 仅用于结构判断的文本规范化：剥离 Markdown 装饰（不影响文档原文）
function paperText(s) {
  return String(s)
    .replace(/^#{1,6}\s*/, '')   // ### 标题
    .replace(/\*\*/g, '')        // **加粗**
    .replace(/\*/g, '')          // *斜体* / 列表 *
    .replace(/^[-+]\s+/, '')     // - 列表项
    .replace(/\`/g, '')          // 行内代码
    .trim();
}

function isRefsHeading(t) {
  return /^(参考文献|references)\s*$/i.test(t);
}

function isAbstractLine(t) {
  return /^(摘\s*要|abstract|key\s*words?|keywords|关键词|关键字)\s*[:：]/i.test(t);
}

function headingLevel(t) {
  if (t.length > 60) return '';
  if (/^(摘\s*要|abstract|key\s*words?|keywords|关键词|关键字|参考文献|references|致谢|谢辞|acknowledg?e?ments?|引言|前言|绪论|结\s*论|conclusion|总结|附录|appendix|目录|contents)\s*$/i.test(t)) return 'h1';
  if (/^第[一二三四五六七八九十百千\d]+[章部分篇]/.test(t)) return 'h1';
  if (/^[一二三四五六七八九十]+、/.test(t)) return 'h1';
  if (/^\d+(\.\d+){2}[、.\s]/.test(t)) return 'h3';
  if (/^\d+(\.\d+)[、.\s]/.test(t)) return 'h2';
  if (/^\d+[、.\s]/.test(t)) return 'h1';
  if (/^（[一二三四五六七八九十]+）/.test(t) || /^\([一二三四五六七八九十]+\)/.test(t)) return 'h2';
  return '';
}

function classifyPaperParagraphs(texts) {
  var roles = new Array(texts.length);
  var first = -1;
  for (var i = 0; i < texts.length; i++) {
    if (texts[i].trim()) { first = i; break; }
  }
  var seenBody = false;   // 是否已进入正文（作者信息区只允许出现在标题之后、正文之前）
  var authorCount = 0;
  var inRefs = false;
  for (var i = 0; i < texts.length; i++) {
    var raw = texts[i].trim();
    var t = paperText(texts[i]);
    if (!t) { roles[i] = 'blank'; continue; }
    if (/^[-=*_]{3,}$/.test(raw)) { roles[i] = 'blank'; continue; } // Markdown 分隔线按空段处理
    if (isRefsHeading(t)) { inRefs = true; roles[i] = 'h1'; seenBody = true; continue; }
    if (inRefs) { roles[i] = 'ref'; continue; }
    if (/^\[\d{1,3}\]/.test(t)) { roles[i] = 'ref'; continue; }
    if (i === first) {
      if (isAbstractLine(t)) { roles[i] = /^(摘\s*要|abstract)/i.test(t) ? 'abstract' : 'keywords'; seenBody = true; }
      else if (headingLevel(t)) { roles[i] = headingLevel(t); seenBody = true; }
      else if (t.length <= 80 && !/[。；！？;!]$/.test(t)) { roles[i] = 'title'; }
      else { roles[i] = 'body'; seenBody = true; }
      continue;
    }
    if (!seenBody && authorCount < 3 && t.length <= 60 && !/[。；！？;!]$/.test(t) && !headingLevel(t) && !isAbstractLine(t)) {
      roles[i] = 'author'; authorCount++; continue;
    }
    if (isAbstractLine(t)) { roles[i] = /^(摘\s*要|abstract)/i.test(t) ? 'abstract' : 'keywords'; seenBody = true; continue; }
    var lvl = headingLevel(t);
    if (lvl) { roles[i] = lvl; seenBody = true; continue; }
    roles[i] = 'body'; seenBody = true;
  }
  return roles;
}

function applyPaperStyle(p, role, stats) {
  if (role === 'blank') return 'skip'; // 空段跳过：无需排版，且部分宿主对空段不返回格式对象
  var spec = PAPER_SPECS[role];
  if (!spec) return 'skip';
  var pf = p.paragraphFormat;
  var f = p.font;
  if (!pf || !f) {
    // 回退：改经整段 Range 的格式代理
    var r = p.getRange();
    if (r) { pf = r.paragraphFormat; f = r.font; }
  }
  if (!pf || !f) return false;
  if (stats && stats[role] !== undefined) stats[role]++;
  f.size = spec.size;
  f.bold = !!spec.bold;
  f.name = spec.east;
  f.nameFarEast = spec.east;        // 中文字体（WordApiDesktop 1.3+）
  f.nameAscii = spec.latin;         // 西文字体
  f.nameOther = spec.latin;
  f.nameBidirectional = spec.latin;
  pf.alignment = spec.center ? 'Centered' : (spec.justify ? 'Justified' : 'Left');
  pf.leftIndent = 0;
  pf.firstLineIndent = 0;
  if (spec.indent) pf.firstLineIndent = spec.indent;                       // 首行缩进（磅：2 字符 × 12 磅 = 24）
  if (spec.hanging) { pf.leftIndent = spec.hanging; pf.firstLineIndent = -spec.hanging; } // 悬挂缩进
  pf.spaceBefore = spec.spaceBefore || 0;
  pf.spaceAfter = spec.spaceAfter || 0;
  pf.lineSpacing = Math.round(spec.size * 1.5);                            // 约 1.5 倍行距（点值）
  if (spec.outline) pf.outlineLevel = spec.outline;
  return 'ok';
}

function paperSummary(stats, ok, blanks, total, fails) {
  var parts = [];
  if (stats.title) parts.push('标题 ' + stats.title);
  if (stats.author) parts.push('作者信息 ' + stats.author);
  if (stats.h1) parts.push('一级标题 ' + stats.h1);
  if (stats.h2) parts.push('二级标题 ' + stats.h2);
  if (stats.h3) parts.push('三级标题 ' + stats.h3);
  if (stats.body) parts.push('正文 ' + stats.body);
  if (stats.ref) parts.push('参考文献 ' + stats.ref);
  var msg = '已完成标准中文论文排版：页面 A4，正文宋体小四、1.5 倍行距、首行缩进 2 字符，标题与各级标题黑体加粗，参考文献五号悬挂缩进。'
    + NL + '识别段落：' + (parts.length ? parts.join(' · ') : '未识别出典型结构，全部按正文处理')
    + '（成功排版 ' + ok + ' 段、空段跳过 ' + blanks + ' 段、失败 ' + fails.length + ' 段，共 ' + total + ' 段）';
  if (fails.length) {
    msg += NL + '失败详情：' + fails.slice(0, 3).join('；') + (fails.length > 3 ? ' 等' : '');
    if (!ok && (total - blanks) > 0) {
      msg += NL + '所有含内容段落都未成功，疑似 Office 版本或文档兼容问题。请把本条信息反馈给开发者。';
    }
  }
  msg += '。可用 Ctrl+Z 撤销。';
  return msg;
}

// 把段落 OOXML 字符串按论文规范改写（不经过格式对象，用于宿主拒绝格式代理时的兜底）
function transformPaperOoxml(xml, role) {
  var spec = PAPER_SPECS[role];
  if (!spec) return null;
  var sz = Math.round(spec.size * 2); // 半磅单位：12pt = 24
  var rfonts = '<w:rFonts w:ascii="' + spec.latin + '" w:hAnsi="' + spec.latin + '" w:eastAsia="' + spec.east + '" w:cs="' + spec.latin + '"/>';
  var sizes = '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>';
  var bold = spec.bold ? '<w:b/>' : '';
  // 1) 清理既有格式标签
  xml = xml.replace(/<w:rFonts[^>]*\/>/g, '');
  xml = xml.replace(/<w:sz[^>]*\/>/g, '');
  xml = xml.replace(/<w:szCs[^>]*\/>/g, '');
  xml = xml.replace(/<w:b[^>]*\/>/g, '');
  xml = xml.replace(/<w:spacing[^>]*\/>/g, '');
  xml = xml.replace(/<w:ind[^>]*\/>/g, '');
  xml = xml.replace(/<w:jc[^>]*\/>/g, '');
  xml = xml.replace(/<w:outlineLvl[^>]*\/>/g, '');
  // 2) 每个 rPr 注入字体/字号/加粗
  xml = xml.replace(/<w:rPr[^>]*>/g, function (tag) {
    return (tag.slice(-2) === '/>') ? '<w:rPr>' + rfonts + sizes + bold + '</w:rPr>' : tag + rfonts + sizes + bold;
  });
  // 3) pPr 注入间距/缩进/对齐/大纲级别
  var jc = spec.center ? 'center' : (spec.justify ? 'both' : 'left');
  var ind = spec.hanging
    ? '<w:ind w:leftChars="200" w:left="420" w:hangingChars="200" w:hanging="420"/>'
    : '<w:ind w:firstLine="' + (spec.indent ? '480' : '0') + '" w:firstLineChars="' + (spec.indent ? '200' : '0') + '" w:left="0"/>';
  var spacing = '<w:spacing w:before="' + (spec.spaceBefore || 0) + '" w:after="' + (spec.spaceAfter || 0) + '" w:line="360" w:lineRule="auto"/>';
  var oLvl = spec.outline ? '<w:outlineLvl w:val="' + spec.outline.replace('OutlineLevel', '') + '"/>' : '';
  var pprAdd = spacing + ind + '<w:jc w:val="' + jc + '"/>' + oLvl;
  if (/<w:pPr\/>/.test(xml)) {
    xml = xml.replace(/<w:pPr\/>/, function () { return '<w:pPr>' + pprAdd + '</w:pPr>'; });
  } else if (/<w:pPr>[\s\S]*?<w:rPr/.test(xml)) {
    // 插到 pPr 内 rPr 之前（保证 pPr 子元素顺序合法）
    xml = xml.replace(/(<w:pPr>[\s\S]*?)<w:rPr/, '$1' + pprAdd + '<w:rPr');
  } else if (/<w:pPr>/.test(xml)) {
    xml = xml.replace(/<w:pPr>/, function () { return '<w:pPr>' + pprAdd; });
  } else {
    xml = xml.replace(/<w:p\b[^>]*>/, function (m) { return m + '<w:pPr>' + pprAdd + '</w:pPr>'; });
  }
  return xml;
}

function formatPaperDocument() {
  // 只读/保护状态前置检查：此时 Word 会拒绝格式写入（段落格式对象返回 undefined）
  if (Office.context && Office.context.document && Office.context.document.mode === Office.DocumentMode.ReadOnly) {
    return Promise.reject(new Error('当前文档处于只读模式（例如从邮件附件直接打开）。请点击文档顶部的「启用编辑」，或把文档另存为新的 .docx 后重试。'));
  }
  var CHUNK = 30; // 每批写入的段落数：分批可避免单批操作过大导致宿主拒绝
  var texts = [];
  var roles = [];
  var stats = { title: 0, author: 0, h1: 0, h2: 0, h3: 0, body: 0, ref: 0 };
  var ok = 0, blanks = 0;
  var fails = [];

  // 阶段一：只读段落文本并分类（不在此批次写任何格式）
  return Word.run(function (context) {
    var paras = context.document.body.paragraphs;
    paras.load('items/text');
    return context.sync().then(function () {
      texts = paras.items.map(function (p) { return p.text; });
      if (!texts.length || !texts.some(function (t) { return t.trim(); })) {
        return '文档为空，没有可排版的内容。';
      }
      roles = classifyPaperParagraphs(texts);
      return null;
    });
  }).then(function (emptyMsg) {
    if (emptyMsg) return emptyMsg;

    // 阶段二：分批写入格式，每批使用全新的集合对象
    function chunk(start) {
      if (start >= texts.length) return Promise.resolve();
      var end = Math.min(start + CHUNK, texts.length);
      return Word.run(function (context) {
        if (start === 0) {
          var page = context.document.sections.getFirst().pageSetup;
          page.paperSize = 'A4';
          page.topMargin = 72; page.bottomMargin = 72;   // 上下 2.54cm
          page.leftMargin = 90; page.rightMargin = 90;   // 左右 3.17cm
        }
        var paras = context.document.body.paragraphs;
        paras.load('items');
        return context.sync().then(function () {
          var fallback = []; // 格式对象不可用、需要 OOXML 兜底的段索引
          for (var i = start; i < end; i++) {
            var p = paras.items[i];
            var role = roles[i];
            if (role === 'blank') { blanks++; continue; }
            try {
              if (role === 'abstract' || role === 'keywords') {
                var m = String(texts[i] || '').match(/^(摘\s*要|abstract|key\s*words?|keywords|关键词|关键字)\s*[:：]/i);
                if (m) p.search(m[1]).getFirst().font.bold = true; // 标签加粗（格式对象可用时生效）
              }
              var rv = applyPaperStyle(p, role, stats);
              if (rv === 'ok') ok++;
              else if (rv === 'skip') blanks++;
              else fallback.push(i);
            } catch (e) {
              fallback.push(i);
            }
          }
          if (!fallback.length) return context.sync();
          // OOXML 兜底：取段落 XML → 字符串改写 → 整段写回（完全不经过格式对象）
          var xmlResults = fallback.map(function (idx) { return paras.items[idx].getOoxml(); });
          return context.sync().then(function () {
            fallback.forEach(function (idx, j) {
              var xml = xmlResults[j] && xmlResults[j].value;
              try {
                var transformed = xml ? transformPaperOoxml(xml, roles[idx]) : null;
                if (transformed) {
                  paras.items[idx].insertOoxml(transformed, Word.InsertLocation.replace);
                  ok++;
                  if (stats[roles[idx]] !== undefined) stats[roles[idx]]++;
                } else {
                  fails.push('第' + (idx + 1) + '段' + (texts[idx] ? '（' + String(texts[idx]).slice(0, 15) + '…）' : '') + '：OOXML 转换失败');
                }
              } catch (e) {
                fails.push('第' + (idx + 1) + '段：OOXML 写入失败：' + (e && e.message ? e.message : e));
              }
            });
            return context.sync();
          });
        });
      }).then(function () { return chunk(end); });
    }
    return chunk(0).then(function () { return paperSummary(stats, ok, blanks, texts.length, fails); });
  });
}

/* ---------------- 动作执行 ---------------- */
function applyLabelFor(action) {
  if (!action.writable) return '';
  if (state.host === 'Excel') {
    return action.writable === 'formula' ? '写入公式到选中单元格' : '写入表格到选中区域';
  }
  return action.writable === 'replace' ? '替换选中文字' : '插入到光标处';
}

function runAction(action, userText) {
  if (state.busy) { toast('正在处理中，请稍候'); return; }
  if (action.requiresPrompt && !userText) {
    toast('请先在输入框写下要求，再点「' + action.label + '」');
    $('#prompt').focus();
    return;
  }
  if (action.run) {
    // 直接执行型动作（如论文排版），不走 AI 管线
    setBusy(true);
    addMessage('user', action.label);
    return Promise.resolve()
      .then(function () { return action.run(); })
      .then(function (summary) { addMessage('assistant', summary || '已完成。'); })
      .catch(function (e) { addMessage('assistant', '出错了：' + (e && e.message ? e.message : e)); })
      .then(function () { setBusy(false); });
  }
  setBusy(true);
  collectContext(action).then(function (ctx) {
    if (action.needsSelection && ctx.emptySelection) {
      addMessage('user', action.label);
      addMessage('assistant', '请先在文档中选中要处理的文字，再点一次。');
      return;
    }
    if (action.needsData) {
      var single = (ctx.rowCount === 1 && ctx.columnCount === 1);
      var blank = single && (ctx.table[0][0] === '' || ctx.table[0][0] === null || ctx.table[0][0] === undefined);
      if (blank) {
        addMessage('user', action.label);
        addMessage('assistant', '请先选中要处理的数据区域（建议包含表头），再点一次。');
        return;
      }
    }
    state.currentAction = action.id;
    addMessage('user', action.label + (userText ? NL + userText : ''));
    var built = action.build(ctx, userText);
    return streamAssistant({
      system: built.system,
      user: built.user,
      writable: action.writable || null,
      applyLabel: applyLabelFor(action),
      maxTokens: action.maxTokens || 4000,
    });
  }).catch(function (e) {
    addMessage('assistant', '出错了：' + (e && e.message ? e.message : e));
  }).then(function () {
    setBusy(false);
  });
}

function freeChat(userText) {
  if (state.busy) { toast('正在处理中，请稍候'); return; }
  if (!state.host) { toast('请用 Word 或 Excel 打开此加载项'); return; }
  setBusy(true);
  var ctxPromise = state.host === 'Word'
    ? Word.run(function (c) {
        var s = c.document.getSelection();
        s.load('text');
        return c.sync().then(function () { return { selected: (s.text || '').trim().slice(0, 6000) }; });
      })
    : Excel.run(function (c) {
        var r = c.workbook.getSelectedRange();
        r.load(['values']);
        return c.sync().then(function () {
          var rows = r.values.slice(0, 50).map(function (x) { return x.slice(0, 30); });
          return { markdown: rowsToMarkdown(rows) };
        });
      });
  ctxPromise.then(function (ctx) {
    addMessage('user', userText);
    var sys;
    if (state.host === 'Word') {
      sys = '你是内嵌在 Word 里的 AI 办公助手，帮用户写作、编辑、回答关于文档的问题。回答用中文，简洁实用。'
        + (ctx.selected ? NL + '用户当前选中的文字：' + NL + ctx.selected : '');
    } else {
      sys = '你是内嵌在 Excel 里的 AI 办公助手，帮用户分析数据、写公式、解决表格问题。回答用中文，简洁实用。'
        + (ctx.markdown && ctx.markdown !== '(空选区)' ? NL + '用户当前选中区域（Markdown，首行为表头）：' + NL + ctx.markdown : '');
    }
    return streamAssistant({ system: sys, user: userText, writable: null });
  }).catch(function (e) {
    addMessage('assistant', '出错了：' + (e && e.message ? e.message : e));
  }).then(function () {
    setBusy(false);
  });
}

/* ---------------- 界面 ---------------- */
function renderActions() {
  var box = $('#actions');
  box.innerHTML = '';
  (ACTIONS[state.host] || []).forEach(function (a) {
    var card = document.createElement('button');
    card.className = 'action-card';
    card.title = a.desc;
    card.innerHTML = '<span class="glyph" aria-hidden="true">&#x' + a.icon + ';</span><div class="txt"><div class="t">' + esc(a.label) + '</div><div class="d">' + esc(a.desc) + '</div></div>';
    card.onclick = function () { onActionClick(a); };
    box.appendChild(card);
  });
}

function onActionClick(action) {
  if (state.busy) { toast('正在处理中，请稍候'); return; }
  var prompt = $('#prompt').value.trim();
  if (action.requiresPrompt && !prompt) {
    toast('请先在输入框写下要求，再点「' + action.label + '」');
    $('#prompt').focus();
    return;
  }
  runAction(action, prompt);
}

function submitPrompt() {
  var text = ($('#prompt').value || '').trim();
  $('#prompt').value = '';
  if (!text) { toast('请输入要求'); return; }
  if (state.busy) { toast('正在处理中，请稍候'); return; }
  if (state.currentAction) runAction(findAction(state.currentAction), text);
  else freeChat(text);
}

function init() {
  loadSettings();
  $('#model').value = state.model;
  $('#temperature').value = String(state.temperature);
  $('#tempVal').textContent = state.temperature;
  $('#chat').innerHTML = '<div class="empty">选择上方一个操作开始，或直接在下方输入框提出要求（回车发送）。' + NL +
    '处理选中内容前：Word 里先选中文字，Excel 里先选中数据区域（建议包含表头）。</div>';
  fetch('/api/config').then(function (r) { return r.json(); }).then(function (c) {
    if (c && c.serverKeySet) $('#serverKeyHint').classList.remove('hidden');
  }).catch(function () {});

  $('#btnSettings').onclick = function () { $('#settingsPanel').classList.toggle('hidden'); };
  $('#model').onchange = function () { state.model = $('#model').value; };
  $('#temperature').oninput = function () { $('#tempVal').textContent = $('#temperature').value; };
  $('#btnSaveSettings').onclick = function () {
    state.model = $('#model').value;
    state.temperature = Number($('#temperature').value);
    saveSettings();
    $('#settingsPanel').classList.add('hidden');
    toast('设置已保存');
  };
  $('#prompt').addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      submitPrompt();
    }
  });
  $('#btnSend').onclick = function () { submitPrompt(); };
}

Office.onReady(function (info) {
  var h = info && info.host ? String(info.host) : '';
  state.host = (h === 'Word') ? 'Word' : (h === 'Excel' ? 'Excel' : null);
  init();
  if (state.host) {
    $('#hostBadge').textContent = state.host + ' 版';
    renderActions();
  } else {
    $('#hostBadge').textContent = '未识别';
    $('#actions').innerHTML = '<div class="empty">请用 Word 或 Excel 打开此加载项（当前宿主：' + esc(h || '未知') + '）</div>';
  }
});
