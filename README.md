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

- 点「在 Word 中打开」——生成 `.docx` 并开始下载，Word 通常会自动打开它（**公式是可双击编辑的原生公式**）；
- 点「复制到 Word」——把内容（含原生公式）写进剪贴板，打开 Word 按 Ctrl+V 直接粘贴，适合不想下载文件时；
- 点「预览」——下方按 A4 纸面显示排版效果，再点一次可收起。

## 预览与复制

预览按 A4 页面（210 × 297 mm）渲染，字体与缩进和导出的 `.docx` 一致，公式由 KaTeX 显示，仅用于查看效果。

「复制到 Word」把内容以带内联样式的 HTML 写进剪贴板（Word 粘贴时不读外部 CSS，所以必须内联），
粘进 Word 后标题、加粗、斜体、列表、表格都会保留，**公式也是可双击编辑的原生公式**。

这里有一个必须注意的实现细节：复制走的是 `navigator.clipboard.write()` 写**字符串原始字节**，
而不是把 DOM 节点交给 `document.execCommand('copy')`。因为浏览器在复制 DOM 时会解析并重新
序列化 HTML，把 OMML 标签名小写化（`<m:oMathPara>` → `<m:omathpara>`）；而 Word 对这几个
标签**大小写敏感**，小写之后公式会被拍平成一行普通文本（分数不再堆叠、上下标全部平掉）。
写原始字节可以完整保留大小写，Word 才会把它建成原生公式。

因此复制功能有两点前提：

（1）页面需在安全上下文打开（双击 `file://` 或 `https`/`localhost` 都满足）；
（2）复制动作由点击触发（浏览器要求用户手势）。

若浏览器不支持或写入被拒绝，会自动降级为 `execCommand('copy')`：文字、表格、样式仍能粘过去，
但公式会退化为文本，此时建议改用「在 Word 中打开」。

另外界面上有「复制提示词」按钮，可把一段要求 AI 输出标准 Markdown 的提示词复制走，
直接贴给大模型，能减少 `\(...\)` 之类不兼容写法。

## 公式

支持的定界符：`$…$`、`$$…$$`、`\(…\)`、`\[…\]`（后两种会自动归一化）。

两条输出路径都给的是 Word 原生 OMML、可双击编辑：

（1）「在 Word 中打开」→ `.docx` 内的 `<m:oMath>`；
（2）「复制到 Word」→ 剪贴板里的 Word-HTML 形态 OMML（条件注释块 + `msEquation`）。

如需 MathType，可在 Word 中用 MathType 插件对已有公式做转换。

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
