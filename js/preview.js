/*
 * preview.js — 把 Markdown 渲染成 HTML。
 * 两个用途，共用同一套 token 遍历，保证「所见 = 所得」：
 *   1. toPreviewHtml(md)            屏幕预览：公式用 KaTeX 渲染，样式走 css/preview.css 的类名
 *   2. toWordHtml(md, tableStyle)   复制到 Word：公式用 OMML，样式全部写成内联
 *                                   （Word 粘贴时不认外部 CSS，必须内联）
 *
 * 依赖全局：marked、katex、MD2DOCX（normalizeMathDelimiters / ensureMarkedMath）、MD2MATH
 * 暴露：global.MD2PREVIEW
 */
(function (global) {
  'use strict';

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  var BODY_FONT = "'Times New Roman','宋体'";
  var HEAD_FONT = "'黑体'";
  var MONO_FONT = "Consolas,'Courier New',monospace";
  var HEAD_SIZE = { h1: '16pt', h2: '14pt', h3: '13pt', h4: '12pt', h5: '11pt', h6: '11pt' };

  // ---- 公式渲染 ----
  function katexMath(tex, display) {
    try {
      return global.katex.renderToString(String(tex || '').trim(), {
        displayMode: !!display,
        output: 'html',
        strict: false,
        throwOnError: false,
      });
    } catch (e) {
      return '<span>' + esc(tex) + '</span>';
    }
  }

  // ---- 复制到 Word：OMML → Word-HTML 形态 ----
  // 背景：Word 的 HTML 导入器不认 docx 形态的 OMML，只认它自己 SaveAs HTML 时的那种形态。
  // 本函数按 Word 的真实输出（用 Word 把含 OMML 的 docx 另存为 HTML 得到，并已验证可回导入为
  // 原生公式）做等价转换。四要点：
  //   1) <m:r><m:rPr>..</m:rPr><m:t>TEXT</m:t></m:r> -> <m:r><m:rPr>..</m:rPr>TEXT</m:r>
  //      （m:t 退化为直接文本，rPr 保留）
  //   2) run 外包一层：无 rPr -> <i style='mso-bidi-font-style:normal'><span lang=EN-US style='…Cambria Math…'>；
  //                     有 rPr -> <span lang=EN-US style='…Cambria Math…'>
  //   3) 仅在 Word 也会输出 ctrlPr 的 *Pr 里注入 <span …><m:ctrlPr></m:ctrlPr></span>；
  //      **绝不能注入 <m:mcPr>**（CT_MCPr 不允许，注入后 Word 的 HTML 解析器直接报错打不开）
  //   4) 剥掉 <m:argPr>/<m:baseJc>（Word 的 HTML 不含它们）与 <m:oMath> 上的 xmlns:m/xmlns:mml
  //      局部声明（让 m: 继承最外层元素上的 2004 omml 命名空间）
  var OMML_NS_2004 = 'http://schemas.microsoft.com/office/2004/12/omml';
  var CAM = 'font-family:"Cambria Math",serif';
  var CAMFULL =
    'font-family:"Cambria Math",serif;mso-ascii-font-family:"Cambria Math";' +
    'mso-hansi-font-family:"Cambria Math"';
  var RE_MRUN = /<m:r>(<m:rPr>[\s\S]*?<\/m:rPr>)?<m:t[^>]*>([\s\S]*?)<\/m:t><\/m:r>/g;
  var RE_PR_CLOSE = /<\/m:(?!rPr\b|mcPr\b|argPr\b)([A-Za-z]+Pr)>/g;
  var RE_RUN_RPR = /<m:r><m:rPr>[\s\S]*?<\/m:rPr>[\s\S]*?<\/m:r>/g;
  var RE_RUN_PLAIN = /<m:r>(?!<m:rPr>)[\s\S]*?<\/m:r>/g;
  var RE_ARG_PR = /<m:argPr>[\s\S]*?<\/m:argPr>/g;
  var RE_BASE_JC = /<m:baseJc[^>]*\/>/g;
  var RE_LOCAL_NS = /\s*xmlns:(m|mml)="[^"]*"/g;

  function ommlToWordHtml(o) {
    o = String(o || '');
    o = o.replace(RE_MRUN, function (m, rPr, txt) {
      return '<m:r>' + (rPr || '') + txt + '</m:r>';
    });
    o = o.replace(RE_PR_CLOSE, function (m, name) {
      return (
        "<span style='" + CAMFULL + "'><m:ctrlPr></m:ctrlPr></span></m:" + name + '>'
      );
    });
    o = o.replace(RE_RUN_RPR, function (m) {
      return "<span lang=EN-US style='" + CAM + "'>" + m + '</span>';
    });
    o = o.replace(RE_RUN_PLAIN, function (m) {
      return (
        "<i style='mso-bidi-font-style:normal'><span lang=EN-US style='" + CAM + "'>" +
        m +
        '</span></i>'
      );
    });
    o = o.replace(RE_ARG_PR, '').replace(RE_BASE_JC, '');
    o = o.replace(RE_LOCAL_NS, '');
    return o;
  }

  // 包成 Word 的条件注释块。
  // 实测结论（2026-10-16，Word 16.0 + 真实剪贴板）：
  //   · Word 的剪贴板 HTML 导入器**认**条件注释里的 OMML，能做成原生公式，
  //     但有两条硬要求：① `m:` 必须绑定到 2004 omml 命名空间；
  //     ② **OMML 标签必须保持原始大小写** —— Word 对 `<m:oMathPara>` 与 `<m:omathpara>`
  //        区别对待，后者会被拍平成一行 Unicode 文本（分数不堆叠、上下标平掉）。
  //   · 因此复制路径**不能**用 document.execCommand('copy')（浏览器会 parse+serialize，
  //     把标签小写化，必然失败），必须用 navigator.clipboard.write 写字符串原始字节。
  //   · 命名空间声明放在公式最外层元素自身即可（实测 oMath=3）；放在最外层 div 上不够。
  // 条件注释保留：Word 直接打开 .html 文件那条路（InsertFile）同样依赖它。
  // 兜底用 LaTeX 而不是空串：引擎不可用/极老 Word 时至少标出公式位置，可粘进 Word 365 的
  // LaTeX 公式输入或 MathType 的 Alt+\ 转换。
  function wrapEquation(ommlHtml, tex, display) {
    var body = ommlHtml;
    if (display && body.indexOf('<m:oMathPara') < 0) {
      body = '<m:oMathPara>' + body + '</m:oMathPara>';
    }
    // 在公式最外层元素上再声明一次命名空间（幂等：已有就不加）
    body = body.replace(/^<m:(oMathPara|oMath)\b([^>]*)>/, function (m, tag, attrs) {
      if (/xmlns:m=/.test(attrs)) return m;
      return '<m:' + tag + ' xmlns:m="' + OMML_NS_2004 + '"' + attrs + '>';
    });
    var fb =
      "<span lang=EN-US style='font-family:Consolas;font-size:9pt'>" + esc(tex) + '</span>';
    return (
      '<!--[if gte msEquation 12]>' + body + '<![endif]-->' +
      '<![if !msEquation]>' + fb + '<![endif]-->'
    );
  }

  function ommlHtmlMath(tex, display) {
    var t = String(tex || '').trim();
    var omml = '';
    try {
      if (global.MD2MATH && typeof global.MD2MATH.mathToOmml === 'function') {
        omml = ommlToWordHtml(global.MD2MATH.mathToOmml(t, !!display));
      }
    } catch (e) {
      omml = '';
    }
    if (!omml) {
      return latexMath(tex, display); // 引擎不可用时退回 LaTeX 原文
    }
    return wrapEquation(omml, t, !!display);
  }

  function latexMath(tex, display) {
    var t = String(tex || '').trim();
    var body = display ? '$$\n' + t + '\n$$' : '$' + t + '$';
    return (
      '<span style="font-family:Consolas,\'Courier New\',monospace;font-size:10.5pt;background:#f2f2f2">' +
      esc(body) +
      '</span>'
    );
  }

  function ommlMath(tex, display) {
    var t = String(tex || '').trim();
    try {
      if (global.MD2MATH && typeof global.MD2MATH.mathToOmml === 'function') {
        return global.MD2MATH.mathToOmml(t, !!display);
      }
    } catch (e) {
      /* 引擎不可用则退回纯文本 */
    }
    return '<span>' + esc(t) + '</span>';
  }

  // ---- 样式片段 ----
  function styleAttr(map) {
    var parts = [];
    for (var k in map) {
      if (Object.prototype.hasOwnProperty.call(map, k) && map[k]) parts.push(k + ':' + map[k]);
    }
    return parts.length ? ' style="' + parts.join(';') + '"' : '';
  }

  function pStyle(inline) {
    return inline ? styleAttr({ margin: '0 0 6pt 0', 'font-family': BODY_FONT, 'font-size': '12pt', 'line-height': '1.5' }) : '';
  }

  function headingTag(t, html, inline) {
    var d = Math.min(Math.max(t.depth || 1, 1), 6);
    var key = 'h' + d;
    if (inline) {
      return (
        '<' + key + styleAttr({ margin: '14pt 0 6pt', 'font-family': HEAD_FONT, 'font-weight': 'bold', 'font-size': HEAD_SIZE[key] }) + '>' +
        html + '</' + key + '>'
      );
    }
    return '<' + key + ' class="md-' + key + '">' + html + '</' + key + '>';
  }

  // ---- 行内 token ----
  function inlineHtml(tokens, math, inline) {
    var out = [];
    for (var i = 0; i < (tokens || []).length; i++) {
      var t = tokens[i];
      switch (t.type) {
        case 'text':
          if (Array.isArray(t.tokens) && t.tokens.length) out.push(inlineHtml(t.tokens, math, inline));
          else out.push(esc(t.text));
          break;
        case 'escape':
          out.push(esc(t.text));
          break;
        case 'strong':
          out.push('<strong>' + inlineHtml(t.tokens, math, inline) + '</strong>');
          break;
        case 'em':
          out.push('<em>' + inlineHtml(t.tokens, math, inline) + '</em>');
          break;
        case 'del':
          out.push('<s>' + inlineHtml(t.tokens, math, inline) + '</s>');
          break;
        case 'codespan':
          out.push(inline ? '<code' + styleAttr({ 'font-family': MONO_FONT, 'font-size': '10.5pt' }) + '>' + esc(t.text) + '</code>' : '<code>' + esc(t.text) + '</code>');
          break;
        case 'br':
          out.push('<br>');
          break;
        case 'link':
          out.push('<a href="' + esc(t.href) + '">' + inlineHtml(t.tokens, math, inline) + '</a>');
          break;
        case 'image':
          out.push('<img src="' + esc(t.href) + '" alt="' + esc(t.text || '') + '">');
          break;
        case 'mathInline':
          out.push(math(t.text, false));
          break;
        default:
          if (t.text) out.push(esc(t.text));
      }
    }
    return out.join('');
  }

  // ---- 块级 token ----
  function listHtml(token, math, inline) {
    var tag = token.ordered ? 'ol' : 'ul';
    var items = [];
    (token.items || []).forEach(function (item) {
      var parts = [];
      (item.tokens || []).forEach(function (t) {
        if (t.type === 'list') {
          parts.push(listHtml(t, math, inline));
        } else if (t.type === 'space' || t.type === 'html') {
          return;
        } else if (t.type === 'text' && Array.isArray(t.tokens) && t.tokens.length) {
          parts.push(inlineHtml(t.tokens, math, inline));
        } else if (t.type === 'paragraph') {
          parts.push(inlineHtml(t.tokens || [], math, inline));
        } else {
          parts.push(inlineHtml([t], math, inline));
        }
      });
      items.push('<li>' + parts.join('') + '</li>');
    });
    return '<' + tag + '>' + items.join('') + '</' + tag + '>';
  }

  // 返回单元格的原始 CSS 串（调用方直接拼 style="…"）
  function cellBorder(three, isHeader, isLastRow) {
    var b = 'padding:4px 8px';
    if (!three) return b + ';border:1px solid #000';
    if (isHeader) b += ';border-top:2pt solid #000;border-bottom:1pt solid #000';
    if (isLastRow) b += ';border-bottom:2pt solid #000';
    return b;
  }

  function tableHtml(token, opts, math, inline) {
    var three = !!inline && opts.tableStyle === 'three-line';
    var rows = (token.rows || []).length;
    function cell(c, isHeader, isLastRow) {
      var runs = inlineHtml((c && c.tokens) || [], math, inline);
      var tag = isHeader ? 'th' : 'td';
      var cls = inline ? '' : ' class="md-cell"';
      var st = inline ? ' style="' + cellBorder(three, isHeader, isLastRow) + '"' : '';
      return '<' + tag + cls + st + '>' + runs + '</' + tag + '>';
    }
    var head = '';
    if (token.header && token.header.length) {
      var hs = [];
      for (var c = 0; c < token.header.length; c++) hs.push(cell(token.header[c], true, false));
      head = '<thead><tr>' + hs.join('') + '</tr></thead>';
    }
    var body = '';
    for (var r = 0; r < rows; r++) {
      var cs = [];
      for (var k = 0; k < (token.rows[r] || []).length; k++) {
        cs.push(cell(token.rows[r][k], false, r === rows - 1));
      }
      body += '<tr>' + cs.join('') + '</tr>';
    }
    var tblStyle = inline
      ? styleAttr({ 'border-collapse': 'collapse', width: '100%', margin: '0 0 8pt 0' })
      : '';
    return '<table class="md-table"' + tblStyle + '>' + head + '<tbody>' + body + '</tbody></table>';
  }

  function blocksHtml(tokens, opts, math, inline) {
    var out = [];
    (tokens || []).forEach(function (t) {
      switch (t.type) {
        case 'heading':
          out.push(headingTag(t, inlineHtml(t.tokens, math, inline), inline));
          break;
        case 'paragraph':
        case 'text':
          out.push('<p' + pStyle(inline) + '>' + inlineHtml(t.tokens || [], math, inline) + '</p>');
          break;
        case 'code':
          out.push(
            inline
              ? '<pre' + styleAttr({ margin: '0 0 8pt 0', 'font-family': MONO_FONT, 'font-size': '10pt', background: '#f2f2f2', padding: '6pt 8pt' }) + '><code>' + esc(t.text || '') + '</code></pre>'
              : '<pre class="md-code"><code>' + esc(t.text || '') + '</code></pre>'
          );
          break;
        case 'blockquote':
          out.push(
            inline
              ? '<blockquote' + styleAttr({ margin: '0 0 8pt 0', 'padding-left': '12pt', 'border-left': '3pt solid #cccccc', color: '#595959' }) + '>' + blocksHtml(t.tokens || [], opts, math, inline) + '</blockquote>'
              : '<blockquote class="md-quote">' + blocksHtml(t.tokens || [], opts, math, inline) + '</blockquote>'
          );
          break;
        case 'list':
          out.push(listHtml(t, math, inline));
          break;
        case 'table':
          out.push(tableHtml(t, opts, math, inline));
          break;
        case 'hr':
          out.push(inline ? '<hr' + styleAttr({ border: 'none', 'border-top': '1pt solid #cccccc' }) + '>' : '<hr>');
          break;
        case 'mathBlock':
          out.push(math(t.text, true));
          break;
        case 'space':
        case 'html':
        default:
          break;
      }
    });
    return out.join('\n');
  }

  // ---- 对外 ----
  function render(md, opts) {
    opts = opts || {};
    var inline = !!opts.inline;
    var math =
      typeof opts.math === 'function'
        ? opts.math
        : opts.math === 'omml-html'
        ? ommlHtmlMath
        : opts.math === 'omml'
        ? ommlMath
        : opts.math === 'latex'
        ? latexMath
        : katexMath;
    var M = global.marked;
    if (!M || typeof M.lexer !== 'function') throw new Error('marked 未加载，请先引入 vendor/marked.min.js');
    if (!global.MD2DOCX || typeof global.MD2DOCX.normalizeMathDelimiters !== 'function') {
      throw new Error('converter.js 未加载或不兼容');
    }
    var normalized = global.MD2DOCX.normalizeMathDelimiters(md || '');
    global.MD2DOCX.ensureMarkedMath(M);
    var tokens = M.lexer(normalized);
    return blocksHtml(tokens, opts, math, inline);
  }

  var API = {
    render: render,
    toPreviewHtml: function (md) {
      return render(md, { inline: false, math: 'katex' });
    },
    // 复制到 Word：公式为 Word-HTML 形态的 OMML（可被 Word 导入为原生可编辑公式），
    // 样式全部内联。最外层 div 带 xmlns:m（Word-HTML 的 2004 omml 命名空间）；
    // 每个公式的最外层元素也各自再声明一次（见 wrapEquation）——实测声明在最外层 div 上
    // 对剪贴板导入不够，必须落在公式元素自身才能被 Word 认成原生公式。
    toWordHtml: function (md, tableStyle) {
      var inner = render(md, { inline: true, math: 'omml-html', tableStyle: tableStyle || 'grid' });
      inner = inner.replace(/<p style="/g, '<p class=MsoNormal style="');
      return (
        '<div xmlns:m="' + OMML_NS_2004 + '" xmlns:w="urn:schemas-microsoft-com:office:word" ' +
        'class=WordSection1>' + inner + '</div>'
      );
    },
    ommlToWordHtml: ommlToWordHtml,
  };
  global.MD2PREVIEW = API;
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof self !== 'undefined' ? self : typeof globalThis !== 'undefined' ? globalThis : this);
