/*
 * app.js — 离线网页前台交互逻辑
 * 依赖全局：marked、ZipUtil、DocxBuilder、MD2DOCX、MD2PREVIEW
 */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const mdInput = $('md-input');
  const fileInput = $('file-input');
  const filenameInput = $('filename');
  const statusEl = $('status');
  const charCount = $('char-count');
  const mathStatus = $('math-status');
  const tableStyleSelect = $('table-style');
  const previewWrap = $('preview-wrap');
  const previewPage = $('preview-page');

  const SAMPLE = `# md2word 示例

这是一个将 **Markdown** 转为 *Word* 文档的演示，包含 \`行内代码\` 与[超链接](https://github.com/)。

> 引用块：本地离线处理，内容不会上传到任何服务器。

## 功能列表

- 标题分级
- 有序与无序列表
  - 支持嵌套
- 代码块与表格

1. 第一步
2. 第二步

\`\`\`python
def greet(name: str) -> str:
    return f"Hello, {name}"
\`\`\`

| 元素 | 支持 | 说明 |
| --- | --- | --- |
| 标题 | ✅ | H1–H6 |
| 表格 | ✅ | 自动加边框 |
| 代码 | ✅ | 等宽字体 |

## 公式示例

行间公式：
$$
E = \\sum_p \\sum_{n_t} \\left[ T_m(p,n_t) - q(p,n_t)\\psi(n_t) \\right]^2
$$

行内公式如 $\\psi(t)$、$\\sigma_T^2$ 也会转为 Word 原生可编辑公式。美元符号成对包裹（单个＝行内，双个＝行间）与反斜杠括号两种写法都支持。

列表里的行内公式同样有效：

- $\\Delta T(z,t)$：深度 $z$ 处、时刻 $t$ 的温升；
- $q$：表面恒定净热流密度；
- $\\lambda$：导热系数；

正文结束。
`;

  function setStatus(msg, type) {
    statusEl.textContent = msg;
    statusEl.className = 'status' + (type ? ' ' + type : '');
  }

  function setMathStatus(msg, cls) {
    if (!mathStatus) return;
    mathStatus.textContent = msg;
    mathStatus.className = 'badge ' + (cls || 'badge-warn');
  }

  // 公式引擎初始化（离线：直接读内嵌的 XSLT 字符串）
  function initMathEngine() {
    if (typeof window.MD2MATH === 'undefined') {
      setMathStatus('公式引擎：不可用（公式将转为文本）', 'badge-err');
      return Promise.resolve(false);
    }
    if (typeof window.katex === 'undefined') {
      setMathStatus('KaTeX 未加载（公式将转为文本）', 'badge-err');
      return Promise.resolve(false);
    }
    setMathStatus('公式引擎：加载中…', 'badge-warn');
    return window.MD2MATH.init()
      .then(function () {
        setMathStatus('公式引擎：就绪 ✓', 'badge-ok');
        return true;
      })
      .catch(function (err) {
        console.error('公式引擎初始化失败:', err);
        setMathStatus('公式引擎：初始化失败（公式将转为文本）', 'badge-err');
        return false;
      });
  }

  function updateCount() {
    charCount.textContent = mdInput.value.length + ' 字符';
  }

  function currentTableStyle() {
    return tableStyleSelect && tableStyleSelect.value === 'three-line' ? 'three-line' : 'grid';
  }

  // 预览：渲染成 Word 风格的 A4 纸面（公式用 KaTeX 渲染，仅供查看）
  function renderPreview() {
    const md = mdInput.value;
    if (!md.trim()) {
      setStatus('请先输入或导入 Markdown 内容', 'err');
      return;
    }
    try {
      previewPage.innerHTML = window.MD2PREVIEW.toPreviewHtml(md);
      previewWrap.hidden = false;
      previewWrap.scrollIntoView({ behavior: 'smooth', block: 'start' });
      setStatus('预览已生成', 'ok');
    } catch (err) {
      console.error(err);
      setStatus('预览失败：' + (err && err.message ? err.message : err), 'err');
    }
  }

  // 收起预览，并把页面滚动回输入区
  function hidePreview() {
    previewWrap.hidden = true;
    setStatus('', '');
    mdInput.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // 预览开关：已显示则收起，未显示则渲染后显示
  function togglePreview() {
    if (!previewWrap.hidden) {
      hidePreview();
      return;
    }
    renderPreview();
  }

  // 把纯文本写入剪贴板（用于复制「给 AI 的提示词」）
  function copyPlainText(text) {
    const holder = document.createElement('div');
    holder.setAttribute('contenteditable', 'true');
    holder.style.position = 'fixed';
    holder.style.left = '-99999px';
    holder.style.top = '0';
    holder.textContent = text; // 用 textContent，避免 HTML 转义干扰
    document.body.appendChild(holder);
    const range = document.createRange();
    range.selectNodeContents(holder);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    let ok = false;
    try {
      ok = document.execCommand('copy');
    } catch (e) {
      ok = false;
    }
    sel.removeAllRanges();
    document.body.removeChild(holder);
    return ok;
  }

  // 降级方案：用 execCommand 复制 DOM 节点。
  // **公式在这条路上会失效**——浏览器读 DOM 时会 parse+serialize，把 OMML 标签名小写化
  //（<m:oMathPara> → <m:omathpara>），而 Word 对标签大小写敏感，公式会被拍平成一行文本。
  // 只在 navigator.clipboard.write 不可用时兜底，至少保证文字/表格/样式能粘过去。
  function copyViaExecCommand(html) {
    const holder = document.createElement('div');
    holder.setAttribute('contenteditable', 'true');
    holder.style.position = 'fixed';
    holder.style.left = '-99999px';
    holder.style.top = '0';
    holder.innerHTML = html;
    document.body.appendChild(holder);
    const range = document.createRange();
    range.selectNodeContents(holder);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    let ok = false;
    try {
      ok = document.execCommand('copy');
    } catch (e) {
      ok = false;
    }
    sel.removeAllRanges();
    document.body.removeChild(holder);
    return ok;
  }

  // ★ 首选方案：把 HTML 作为「字符串原始字节」写进剪贴板。
  // 这条路不经过 HTML 解析器，OMML 标签的大小写原样保留，Word 粘贴时能建成原生可编辑公式。
  // 需要安全上下文（file:// 在 Chrome 里也算）且由用户手势触发；不可用时返回 false 走降级。
  function copyRawHtml(html, plain) {
    if (typeof navigator === 'undefined' || !navigator.clipboard) return Promise.resolve(false);
    if (typeof ClipboardItem === 'undefined') return Promise.resolve(false);
    try {
      const item = new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([plain], { type: 'text/plain' }),
      });
      return navigator.clipboard.write([item]).then(
        function () {
          return true;
        },
        function () {
          return false;
        }
      );
    } catch (e) {
      return Promise.resolve(false);
    }
  }

  function htmlToPlain(html) {
    return String(html)
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/<[^>]+>/g, '')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  // 复制到 Word：公式为 Word-HTML 形态 OMML、样式全部内联（Word 粘贴时不认外部 CSS）
  function copyForWord() {
    const md = mdInput.value;
    if (!md.trim()) {
      setStatus('请先输入或导入 Markdown 内容', 'err');
      return;
    }
    const btn = $('btn-copy');
    btn.disabled = true;
    setStatus('正在生成可复制的 Word 内容…', '');
    initMathEngine()
      .then(function () {
        const html = window.MD2PREVIEW.toWordHtml(md, currentTableStyle());
        return copyRawHtml(html, htmlToPlain(html)).then(function (rawOk) {
          if (rawOk) {
            setStatus('已复制：文字、表格、样式与公式都能直接用（公式是 Word 原生可编辑公式）', 'ok');
          } else if (copyViaExecCommand(html)) {
            setStatus(
              '已复制，但当前浏览器不允许写入原生公式，公式会退化成文本；' +
                '需要可双击编辑的公式请用「在 Word 中打开」',
              'err'
            );
          } else {
            setStatus('复制失败，请改用「在 Word 中打开」', 'err');
          }
        });
      })
      .catch(function (err) {
        console.error(err);
        setStatus('复制失败：' + (err && err.message ? err.message : err), 'err');
      })
      .then(function () {
        btn.disabled = false;
      });
  }

  function sanitizeFilename(name) {
    let n = (name || '').trim().replace(/[\\/:*?"<>|]+/g, '_');
    if (!n) n = 'document';
    return n;
  }

  function downloadDocx(uint8, filename) {
    const blob = new Blob([uint8], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // 选择文件
  $('btn-file').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      mdInput.value = reader.result;
      updateCount();
      setStatus('已载入文件：' + file.name, 'ok');
    };
    reader.onerror = () => setStatus('文件读取失败', 'err');
    reader.readAsText(file);
    fileInput.value = '';
  });

  // 示例
  $('btn-sample').addEventListener('click', () => {
    mdInput.value = SAMPLE;
    updateCount();
    setStatus('已填入示例内容', 'ok');
  });

  // 清空
  $('btn-clear').addEventListener('click', () => {
    mdInput.value = '';
    updateCount();
    setStatus('');
  });

  // 转换并下载（Word 通常会自动打开下载到的 .docx）
  $('btn-convert').addEventListener('click', () => {
    const md = mdInput.value;
    if (!md.trim()) {
      setStatus('请先输入或导入 Markdown 内容', 'err');
      return;
    }
    const btn = $('btn-convert');
    btn.disabled = true;
    setStatus('正在转换…', '');
    // 等待公式引擎就绪（若不可用则降级为文本，不阻断转换）
    initMathEngine().then(function () {
      try {
        const uint8 = MD2DOCX.convertToDocx(md, { tableStyle: currentTableStyle() });
        const fname = sanitizeFilename(filenameInput.value) + '.docx';
        downloadDocx(uint8, fname);
        setStatus(
          '已生成 ' + fname + '（' + uint8.length + ' 字节），Word 通常会自动打开；' +
            '若没有，到浏览器下载列表点一下即可。公式在 Word 里是可双击编辑的原生公式',
          'ok'
        );
      } catch (err) {
        console.error(err);
        setStatus('转换失败：' + (err && err.message ? err.message : err), 'err');
      } finally {
        btn.disabled = false;
      }
    });
  });

  // 预览（可收起）
  $('btn-preview').addEventListener('click', togglePreview);
  $('btn-preview-close').addEventListener('click', hidePreview);

  // 复制到 Word
  $('btn-copy').addEventListener('click', copyForWord);

  // 复制「给 AI 的提示词」
  $('btn-copy-prompt').addEventListener('click', () => {
    const text = $('prompt-text').textContent.trim();
    setStatus(copyPlainText(text) ? '提示词已复制，可粘贴给 AI' : '复制失败，请手动选中复制', 'ok');
  });

  mdInput.addEventListener('input', updateCount);
  updateCount();
})();
