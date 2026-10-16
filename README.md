# md2word

把 Markdown 转成 `.docx` 的纯前端离线工具：解析与生成都在浏览器本地完成，文件不上传、不联网。

## 特性

- 完全离线，无后端、无网络请求、无 npm 依赖、无构建步骤
- 支持标题、加粗 / 斜体 / 行内代码、有序与无序列表（含嵌套）、代码块、表格、引用块、分割线、超链接
- 表格可选普通表格（全框线）或三线表（顶线 + 表头下线 + 底线，无竖线）
- 数学公式导出为 Word 原生可编辑公式（OMML），不是图片也不是文本
- A4 纸面预览，所见即所得；也可把内容直接复制后粘贴进 Word
- 中文字体：正文宋体、标题黑体、代码 Consolas

## 使用方法

双击 `md2word.html` 用浏览器打开即可。

若浏览器对本地文件有限制，起一个静态服务：

```bash
python -m http.server 8080
# 访问 http://localhost:8080
```

步骤：左侧粘贴 Markdown 或点「选择文件」导入 `.md` / `.txt`，右侧填导出文件名、切换表格样式，然后：

- 点「转换并下载 Word」——得到 `.docx` 文件；
- 点「预览」——下方按 A4 纸面显示排版效果，再点一次可收起；
- 点「复制到 Word」——把内容写进剪贴板，打开 Word 按 Ctrl+V 直接粘贴。

## 预览与复制

预览按 A4 页面（210 × 297 mm）渲染，字体与缩进和导出的 `.docx` 一致，公式由 KaTeX 显示，仅用于查看效果。

「复制到 Word」生成带内联样式的 HTML 写进剪贴板（Word 粘贴时不读外部 CSS，所以必须内联），
粘进 Word 后标题、加粗、斜体、列表、表格都会保留。

公式同样会保留为 **Word 原生可编辑公式（OMML）**，双击即可编辑。

这里有个坑值得记录：Word 的 HTML 导入器**不认 docx 形态的 OMML**。直接把 `<m:oMath>` 塞进
复制的 HTML，Word 会把它拍平成纯文本（`θ=[A1A2⋯AJσxyσz]⊤∈RJ+2`）。正确做法是先把 OMML 转成
**Word 自己 SaveAs HTML 时用的那种形态**，要点有四：

1. `<m:r><m:rPr>…</m:rPr><m:t>TEXT</m:t></m:r>` → `<m:r><m:rPr>…</m:rPr>TEXT</m:r>`（`<m:t>` 退化为直接文本）；
2. run 外要包 `<i style='mso-bidi-font-style:normal'><span lang=EN-US style='font-family:"Cambria Math",serif'>`；
3. `<m:ctrlPr>` 只注入到 Word 也会输出的 `*Pr` 里——**绝不能注入 `<m:mcPr>`**，否则 Word 的 HTML 解析器会直接报错、打不开文件；
4. 剥掉 `<m:argPr>`、`<m:baseJc>`，并去掉 `<m:oMath>` 上的 `xmlns:m` 局部声明，让 `m:` 继承最外层元素上的 `http://schemas.microsoft.com/office/2004/12/omml`。

公式整体还要包在 `<!--[if gte msEquation 12]> … <![endif]-->` 条件注释里。
以上结论在 Word 16.0（Office 2019/2021，Build 17932）上实测：8 个公式（含矩阵、上下标）
全部导入为原生 OMML，PDF 中可检出 CambriaMath 字体。

另外界面上有「复制提示词」按钮，可把一段要求 AI 输出标准 Markdown 的提示词复制走，
直接贴给大模型，能减少 `\(...\)` 之类不兼容写法。

## 公式

支持的定界符：`$…$`、`$$…$$`、`\(…\)`、`\[…\]`。公式导出为 Word 原生 OMML；如需 MathType，可在 Word 中用 MathType 插件转换。

## 项目结构

```
md2word.html          页面入口
css/style.css         界面样式
css/preview.css       A4 纸面预览样式
js/app.js             前台交互
js/converter.js       Markdown → OOXML
js/docx-builder.js    DOCX 部件组装
js/math.js            LaTeX → MathML → OMML
js/preview.js         Markdown → HTML（屏幕预览 / 复制到 Word）
js/zip.js             ZIP 写入器
vendor/               marked、KaTeX（含 CSS 与字体）、MML2OMML.XSL（本地副本）
LICENSE
```

## 技术栈

marked 解析 Markdown，KaTeX 输出 MathML，微软 `MML2OMML.XSL` 经浏览器 `XSLTProcessor` 转成 OMML，由自研 OOXML 构建器与极简 ZIP 写入器打包成 `.docx`。无构建步骤，无外部依赖。

## License

MIT，见 [LICENSE](LICENSE)。
