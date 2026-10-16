/*
 * converter.js — Markdown -> DOCX 转换器
 * 解析：依赖全局 marked（vendor/marked.min.js）
 * 生成：依赖 docx-builder.js（全局 DocxBuilder） / zip.js（全局 ZipUtil）
 * 浏览器与 Node 通用。
 */
(function (global) {
  'use strict';

  function getBuilder() {
    if (typeof module !== 'undefined' && module.exports) {
      return require('./docx-builder.js');
    }
    return global.DocxBuilder;
  }
  function getMarked() {
    return globalThis.marked;
  }

  // 把「代码」摘出来，不参与公式定界符归一化：
  //   - fenced code block（```/~~~，含未闭合者，保护到文末）
  //   - 行内代码段 `...`
  // 否则贴一段 LaTeX / Shell / JSON 源码时，里面的 \(...\) \[...\]
  // 会被当成公式改写，把代码块内容和行数都改坏。
  function protectCode(text) {
    const stash = [];
    const put = (s) => {
      stash.push(s);
      return '\u0000' + (stash.length - 1) + '\u0000';
    };
    const out = [];
    let fence = null;
    let fenceLen = 0;
    let buf = [];
    for (const line of text.split('\n')) {
      const m = /^ {0,3}(`{3,}|~{3,})/.exec(line);
      if (fence === null) {
        if (m) {
          fence = m[1][0];
          fenceLen = m[1].length;
          buf = [line];
        } else {
          out.push(line.replace(/`[^`\n]*`/g, put));
        }
        continue;
      }
      buf.push(line);
      const closing =
        m && m[1][0] === fence && m[1].length >= fenceLen && /^ {0,3}[`~]+[ \t]*$/.test(line);
      if (closing) {
        out.push(put(buf.join('\n')));
        fence = null;
        fenceLen = 0;
        buf = [];
      }
    }
    if (fence !== null) out.push(put(buf.join('\n'))); // 未闭合围栏：保护到文末
    return {
      text: out.join('\n'),
      restore: function (s) {
        return s.replace(/\u0000(\d+)\u0000/g, function (_, i) {
          return stash[Number(i)];
        });
      },
    };
  }

  // 把 LaTeX 原生定界符统一成 $ 风格（规避 marked 把 \( 当转义、且 pandoc/浏览器对原生定界符支持不一致的问题）
  function normalizeMathDelimiters(md) {
    const p = protectCode(String(md || ''));
    const fixed = p.text
      .replace(/\\\[([\s\S]+?)\\\]/g, function (_, g) {
        return '\n\n$$' + g + '$$\n\n';
      })
      .replace(/\\\(([\s\S]+?)\\\)/g, function (_, g) {
        return '$' + g + '$';
      });
    return p.restore(fixed);
  }

  // 找 needle 在 src 中「不在行内代码段内」的下一个位置；找不到返回 -1。
  // marked 会用 start() 的返回值截断待匹配文本，若把代码段里的 $ 也算进去，
  // 会把 `$x$` 这类代码段截断，导致美元符号被当成公式解析。
  function startOutsideCode(src, needle) {
    let inCode = false;
    for (let i = 0; i < src.length; i++) {
      const c = src[i];
      if (c === '`') inCode = !inCode;
      else if (!inCode && src.startsWith(needle, i)) return i;
    }
    return -1;
  }

  // ---- marked 数学扩展：块级 $$...$$，行内 $...$ ----
  let markedConfigured = false;
  const mathBlockExt = {
    name: 'mathBlock',
    level: 'block',
    start: function (src) {
      const i = startOutsideCode(src, '$$');
      return i < 0 ? undefined : i;
    },
    tokenizer: function (src) {
      const m = /^\$\$([\s\S]+?)\$\$/.exec(src);
      if (m && m[1].indexOf('`') < 0) {
        return { type: 'mathBlock', raw: m[0], text: m[1].trim(), display: true };
      }
      return undefined;
    },
    renderer: function () {
      return '';
    },
  };
  const mathInlineExt = {
    name: 'mathInline',
    level: 'inline',
    start: function (src) {
      const i = startOutsideCode(src, '$');
      return i < 0 ? undefined : i;
    },
    tokenizer: function (src) {
      const m = /^\$([^$\s](?:[^$\n]*?[^$\s])?)\$/.exec(src);
      if (m && m[1].indexOf('`') < 0) {
        return { type: 'mathInline', raw: m[0], text: m[1].trim(), display: false };
      }
      return undefined;
    },
    renderer: function () {
      return '';
    },
  };
  function ensureMarkedMath(M) {
    if (markedConfigured || !M || typeof M.use !== 'function') return;
    M.use({ extensions: [mathBlockExt, mathInlineExt] });
    markedConfigured = true;
  }

  // 超链接注册：url -> rId（rId2, rId3...），顺序与 document.xml.rels 一致
  let hyperlinks = [];
  const linkMap = new Map();
  function getHyperlinkRId(url) {
    if (linkMap.has(url)) return linkMap.get(url);
    hyperlinks.push(url);
    const rid = 'rId' + (2 + hyperlinks.length - 1);
    linkMap.set(url, rid);
    return rid;
  }

  function makeRun(text, style) {
    const rPr = [];
    if (style.bold) rPr.push('<w:b/>');
    if (style.italic) rPr.push('<w:i/>');
    if (style.strike) rPr.push('<w:strike/>');
    if (style.code)
      rPr.push(
        '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas" w:cs="Consolas"/>'
      );
    if (style.link) rPr.push('<w:color w:val="0563C1"/><w:u w:val="single"/>');
    const rPrXml = rPr.length ? `<w:rPr>${rPr.join('')}</w:rPr>` : '';
    return `<w:r>${rPrXml}<w:t xml:space="preserve">${escapeXML(
      text == null ? '' : text
    )}</w:t></w:r>`;
  }

  function escapeXML(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function runsFromTokens(tokens, style) {
    const out = [];
    if (!tokens) return out;
    for (const t of tokens) {
      switch (t.type) {
        // marked 里「块级 text」（列表项、引用内的裸段落）会把行内子 token 挂在
        // t.tokens 上；「行内 text」才是叶子节点。早先直接把 t.text 当纯文本输出，
        // 于是列表项里的 $…$ /**…**/ `…` 全部原样落到 Word（行内公式因此显示成
        // 字面量 "$\Delta T(z,t)$"）。这里必须先递归 t.tokens。
        case 'text':
          if (Array.isArray(t.tokens) && t.tokens.length) {
            out.push(...runsFromTokens(t.tokens, style));
          } else {
            out.push(makeRun(t.text, style));
          }
          break;
        // 松散列表项 / 引用里可能是 paragraph 包裹
        case 'paragraph':
          out.push(...runsFromTokens(t.tokens || [], style));
          break;
        case 'escape':
          out.push(makeRun(t.text, style));
          break;
        case 'strong':
          out.push(...runsFromTokens(t.tokens, { ...style, bold: true }));
          break;
        case 'em':
          out.push(...runsFromTokens(t.tokens, { ...style, italic: true }));
          break;
        case 'codespan':
          out.push(makeRun(t.text, { ...style, code: true }));
          break;
        case 'del':
          out.push(...runsFromTokens(t.tokens, { ...style, strike: true }));
          break;
        case 'br':
          out.push('<w:r><w:br/></w:r>');
          break;
        case 'link': {
          const rid = getHyperlinkRId(t.href);
          const inner = runsFromTokens(t.tokens, { ...style, link: true });
          out.push(
            `<w:hyperlink r:id="${rid}" w:history="1">${inner.join('')}</w:hyperlink>`
          );
          break;
        }
        case 'image':
          out.push(makeRun('[图片: ' + (t.text || t.title || '') + ']', style));
          break;
        case 'mathInline': {
          const tex = (t.text || '').trim();
          try {
            const omml = globalThis.MD2MATH.mathToOmml(tex, false);
            // 行内公式是 run-level 元素，必须是 <w:p> 的直接子级，不能包进 <w:r>
            // （w:r 的内容模型里没有 m:oMath，Word 会按非法嵌套处理）。
            out.push(omml);
          } catch (e) {
            out.push(makeRun('[公式:' + tex + ']', style));
          }
          break;
        }
        default:
          if (t.text) out.push(makeRun(t.text, style));
      }
    }
    return out;
  }

  function listToXml(token, ctx) {
    const out = [];
    const ordered = !!token.ordered;
    const start = Number(token.start) || 1;
    const depth = (ctx && ctx.depth) || 0;
    const indent = 720 + depth * 360;

    token.items.forEach((item, idx) => {
      const blocks = [];
      const nestedLists = [];
      for (const t of item.tokens || []) {
        if (t.type === 'list') nestedLists.push(t);
        else if (t.type === 'space' || t.type === 'html') continue;
        else blocks.push(t);
      }
      const prefix = ordered ? `${start + idx}.  ` : '•  ';
      if (!blocks.length) blocks.push({ type: 'paragraph', tokens: [] });

      // 第一个块与项目符号同一段；其余块（松散列表的第二段等）各成一段并缩进对齐。
      blocks.forEach((b, bi) => {
        const isInlineBlock = b.type === 'paragraph' || b.type === 'text';
        if (isInlineBlock) {
          const innerTokens = Array.isArray(b.tokens)
            ? b.tokens
            : [{ type: 'text', text: b.text || '' }];
          const runs = runsFromTokens(innerTokens, {});
          const head = bi === 0 ? makeRun(prefix, {}) : '';
          const hanging = bi === 0 ? ' w:hanging="360"' : '';
          out.push(
            `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:ind w:left="${indent}"${hanging}/></w:pPr>${head}${runs.join(
              ''
            )}</w:p>`
          );
        } else {
          for (const p of blockToXml(b, ctx)) out.push(p);
        }
      });

      for (const n of nestedLists) {
        out.push(...listToXml(n, { ...ctx, depth: depth + 1 }));
      }
    });
    return out;
  }

  // threeLine：三线表（顶线 + 表头下线 + 底线，无竖线与其余横线）
  function rowToXml(cells, isHeader, threeLine) {
    let xml = isHeader
      ? '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
      : '<w:tr>';
    for (const cell of cells) {
      const cellTokens = cell && cell.tokens ? cell.tokens : [];
      const runs = runsFromTokens(cellTokens, isHeader ? { bold: true } : {});
      // 三线表：表头单元格单独加一条底线（表头下线）；tcW 在前、tcBorders 在后
      const cellBorders =
        threeLine && isHeader
          ? '<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/></w:tcBorders>'
          : '';
      xml += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>${cellBorders}</w:tcPr><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>${runs.join(
        ''
      )}</w:p></w:tc>`;
    }
    xml += '</w:tr>';
    return xml;
  }

  /**
   * 表格 -> OOXML。
   * @param {Object} token marked 的 table token
   * @param {string} [tableStyle] 'grid'（普通表格，全框线，默认）或 'three-line'（三线表）
   */
  function tableToXml(token, tableStyle) {
    const colCount = token.header.length || 1;
    const threeLine = tableStyle === 'three-line';
    const borders = threeLine
      ? `<w:tblBorders>
      <w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>
      <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>
      <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tblBorders>`
      : `<w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>`;
    // 三线表不挂 tblStyle（边框已逐条写明）；普通表格沿用 TableGrid 样式
    const tblPr = threeLine
      ? `<w:tblPr><w:tblW w:w="5000" w:type="pct"/>${borders}</w:tblPr>`
      : `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>${borders}</w:tblPr>`;
    let xml = `<w:tbl>${tblPr}<w:tblGrid>`;
    for (let i = 0; i < colCount; i++) {
      xml += `<w:gridCol w:w="${Math.floor(5000 / colCount)}"/>`;
    }
    xml += '</w:tblGrid>';
    xml += rowToXml(token.header, true, threeLine);
    for (const row of token.rows) xml += rowToXml(row, false, threeLine);
    xml += '</w:tbl>';
    return xml;
  }

  function blockToXml(token, ctx) {
    switch (token.type) {
      case 'heading': {
        const runs = runsFromTokens(token.tokens, {});
        const depth = Math.min(Math.max(token.depth || 1, 1), 6);
        return [
          `<w:p><w:pPr><w:pStyle w:val="Heading${depth}"/></w:pPr>${runs.join(
            ''
          )}</w:p>`,
        ];
      }
      case 'paragraph':
      case 'text': {
        const runs = runsFromTokens(token.tokens || [], {});
        const pStyle = ctx && ctx.quote ? 'Quote' : 'Normal';
        return [
          `<w:p><w:pPr><w:pStyle w:val="${pStyle}"/></w:pPr>${runs.join(
            ''
          )}</w:p>`,
        ];
      }
      case 'code': {
        const lines = (token.text || '').replace(/\n$/, '').split('\n');
        const runs = lines
          .map((ln, i) => {
            const r =
              `<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` +
              `<w:t xml:space="preserve">${escapeXML(ln.length ? ln : ' ')}</w:t></w:r>`;
            return r + (i < lines.length - 1 ? '<w:r><w:br/></w:r>' : '');
          })
          .join('');
        return [
          `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>${runs}</w:p>`,
        ];
      }
      case 'blockquote': {
        const inner = [];
        for (const t of token.tokens || []) {
          inner.push(...blockToXml(t, { ...ctx, quote: true }));
        }
        return inner;
      }
      case 'list':
        return listToXml(token, ctx);
      case 'table':
        return [tableToXml(token, ctx && ctx.tableStyle)];
      case 'hr':
        return [
          `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`,
        ];
      case 'mathBlock': {
        const tex = (token.text || '').trim();
        try {
          const omml = globalThis.MD2MATH.mathToOmml(tex, true);
          return [
            `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>${omml}</w:p>`,
          ];
        } catch (e) {
          return [
            `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>${makeRun(
              '[公式转换失败: ' + tex + ']',
              {}
            )}</w:p>`,
          ];
        }
      }
      case 'space':
      case 'html':
      default:
        return [];
    }
  }

  /**
   * 将 Markdown 文本转换为 .docx 二进制（Uint8Array）。
   * @param {string} markdown
   * @param {Object} [options]
   * @param {string} [options.tableStyle] 'grid'（普通表格，默认）或 'three-line'（三线表）
   * @returns {Uint8Array}
   */
  function convertToDocx(markdown, options) {
    const tableStyle =
      options && options.tableStyle === 'three-line' ? 'three-line' : 'grid';
    const M = getMarked();
    if (!M || typeof M.lexer !== 'function') {
      throw new Error('marked 未加载，请先引入 vendor/marked.min.js');
    }
    const Builder = getBuilder();
    hyperlinks = [];
    linkMap.clear();

    const normalized = normalizeMathDelimiters(markdown || '');
    ensureMarkedMath(M);
    const tokens = M.lexer(normalized);
    const bodyParts = [];
    for (const t of tokens) {
      const parts = blockToXml(t, { tableStyle });
      for (const p of parts) bodyParts.push(p);
    }
    const bodyXml = bodyParts.join('\n');
    return Builder.buildDocx(bodyXml, hyperlinks);
  }

  const API = { convertToDocx, escapeXML, normalizeMathDelimiters, ensureMarkedMath };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else global.MD2DOCX = API;
})(typeof self !== 'undefined' ? self : this);
