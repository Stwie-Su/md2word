/*
 * math.js — 离线公式引擎（浏览器端）
 * 流程：LaTeX 文本 -> KaTeX 解析为 MathML -> 微软 MML2OMML.XSL (XSLT 1.0) 转为 Word 原生 OMML。
 * 依赖：
 *   - vendor/katex.min.js  （全局 katex）
 *   - vendor/mml2omml.xsl.js （全局 window.__MML2OMML_XSL__，离线内嵌；file:// 双击也能用）
 * 暴露：global.MD2MATH = { init, mathToOmml, isReady }
 */
(function (global) {
  'use strict';

  let xsltProcessor = null;
  let ready = false;
  let initPromise = null;

  function getEmbeddedXslt() {
    return global.window && global.window.__MML2OMML_XSL__
      ? global.window.__MML2OMML_XSL__
      : null;
  }

  function buildProcessor(xsltText) {
    const xsltDoc = new DOMParser().parseFromString(xsltText, 'application/xml');
    if (xsltDoc.querySelector('parsererror')) {
      throw new Error('MML2OMML.XSL 解析失败');
    }
    const proc = new XSLTProcessor();
    proc.importStylesheet(xsltDoc);
    return proc;
  }

  // 初始化公式引擎（异步）。成功 resolve，失败 reject。
  function init() {
    if (initPromise) return initPromise;
    initPromise = new Promise((resolve, reject) => {
      const embedded = getEmbeddedXslt();
      if (embedded) {
        try {
          xsltProcessor = buildProcessor(embedded);
          ready = true;
          resolve();
        } catch (e) {
          reject(e);
        }
        return;
      }
      // 回退：通过 http 服务加载（离线双击 file:// 时应已内嵌，不会走到这里）
      fetch('vendor/mml2omml.xsl')
        .then(function (r) {
          if (!r.ok) throw new Error('HTTP ' + r.status);
          return r.text();
        })
        .then(function (t) {
          xsltProcessor = buildProcessor(t);
          ready = true;
          resolve();
        })
        .catch(reject);
    });
    return initPromise;
  }

  // 将 LaTeX 文本渲染为 MathML 的 <math> 元素
  // display：是否为行间公式。必须如实传给 KaTeX —— \tag / \displaystyle / \limits 等
  // 仅在特定模式下合法，传错会导致 KaTeX 抛 “\tag works only in display mode” 之类的错误，
  // 进而整条公式降级成裸文本。
  function renderMathml(tex, display) {
    const html = global.katex.renderToString(tex, {
      output: 'mathml',
      strict: false,
      throwOnError: false,
      displayMode: !!display,
    });
    const div = document.createElement('div');
    div.innerHTML = html;
    return div.querySelector('math');
  }

  function texToMathmlElement(tex, display) {
    const katex = global.katex;
    if (!katex) throw new Error('KaTeX 未加载');
    let mathEl = renderMathml(tex, display);
    // 有些 LaTeX 构造（\tag、\displaystyle、\limits 等）只在 display 模式合法。
    // 公式若被当作行内（display=false）解析却含这类构造，KaTeX 在 mathml 输出下
    // 不会产出 <math> 元素，导致下游转换抛 “KaTeX 未产出 <math> 元素” 而整条公式
    // 退化成红字裸文本。这里自动用相反模式重试一次，优先救活公式。
    if (!mathEl && !display) mathEl = renderMathml(tex, true);
    if (!mathEl) throw new Error('KaTeX 未产出 <math> 元素: ' + tex);
    return mathEl;
  }

  // MathML 元素 -> OMML 字符串
  function mathmlToOmml(mathEl) {
    if (!xsltProcessor) throw new Error('XSLT 未初始化');
    const frag = xsltProcessor.transformToFragment(mathEl, document);
    return new XMLSerializer().serializeToString(frag);
  }

  /**
   * LaTeX -> OMML
   * @param {string} tex  LaTeX 源码（不含定界符 $ $ / \( \)）
   * @param {boolean} display 是否行间公式（true 包 oMathPara，false 行内 m:oMath）
   * @returns {string} OMML XML 片段
   */
  const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

  // 微软 XSLT 把 n-ary（∑/∫ 等）的操作数放在运算符之外（兄弟节点），并生成空的 <m:e/>；
  // Word 对空的操作数基座会渲染成占位方块 □。这里把紧随其后的操作数搬进 <m:e>，
  // 与 pandoc 的结构一致，从而消除占位框。逆序遍历以支持 ∑∑ 连续链。
  function restructureNary(doc) {
    const list = doc.getElementsByTagNameNS(MATH_NS, 'nary');
    const naries = [];
    for (let i = 0; i < list.length; i++) naries.push(list[i]);
    for (let i = naries.length - 1; i >= 0; i--) {
      const nary = naries[i];
      let e = null;
      for (let c = nary.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && c.localName === 'e' && c.namespaceURI === MATH_NS) {
          e = c;
          break;
        }
      }
      const isEmpty = !e || !e.firstChild;
      if (!isEmpty) continue;
      let sib = nary.nextSibling;
      while (sib && sib.nodeType !== 1) sib = sib.nextSibling;
      if (!sib) continue; // 无操作数（罕见），保持原样
      if (!e) {
        e = doc.createElementNS(MATH_NS, 'm:e');
        nary.appendChild(e);
      }
      e.appendChild(sib); // 移动兄弟节点进入基座
    }
  }

  function cleanOmml(omml) {
    const doc = new DOMParser().parseFromString(String(omml), 'application/xml');
    restructureNary(doc);
    return new XMLSerializer().serializeToString(doc.documentElement);
  }

  function mathToOmml(tex, display) {
    const mathEl = texToMathmlElement(String(tex || '').trim(), display);
    let omml = cleanOmml(mathmlToOmml(mathEl));
    if (display && !/<m:oMathPara[ >]/.test(omml)) {
      omml = '<m:oMathPara>' + omml + '</m:oMathPara>';
    }
    return omml;
  }

  const API = {
    init: init,
    mathToOmml: mathToOmml,
    cleanOmml: cleanOmml,
    isReady: function () {
      return ready;
    },
  };
  global.MD2MATH = API;
})(typeof self !== 'undefined' ? self : this);
