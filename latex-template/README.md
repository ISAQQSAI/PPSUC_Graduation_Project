# 信网学院本科毕业设计 LaTeX 模板

这份模板根据学校原始 Word 模板整理而成，当前已经可以稳定完成：

- `LaTeX -> PDF`
- `LaTeX -> DOCX`

其中：

- `PDF` 使用“官方固定样张页 + LaTeX 可编辑正文页”的混合方案
- `DOCX` 使用“官方 Word 模板 + 从 main.tex 自动同步内容”的方案

## 文件说明

- `main.tex`：主模板，日常只需要编辑这个文件
- `main.pdf`：当前编译输出 PDF
- `main.docx`：当前导出 Word 文件
- `assets/`：模板图片、样张 PDF 等资源
- `fonts/`：模板依赖的中文字体文件

## 编译 PDF

建议在当前目录执行：

```bash
latexmk -xelatex -interaction=nonstopmode -halt-on-error main.tex
```

或者：

```bash
xelatex main.tex
xelatex main.tex
```

## 导出 Word

回到仓库根目录执行：

```bash
cd ..
python3 scripts/export_word.py
```

默认输出：

```text
latex-template/main.docx
```

默认模式会直接复用官方 `docx` 模板并替换占位内容，所以版式会更贴近学校模板。

现在它会同步 `main.tex` 里的这些内容：

- 论文基础信息
- 摘要与关键词
- 正文章节标题和正文段落
- 原生 Word 公式（OMML）
- 表格、图片、图题
- 代码块
- 结论、致谢
- 参考文献
- 附录标题和说明

如果你想单独控制 Word 里的结论、致谢、参考文献和附录内容，可填写对应的 `\conclusiontext`、`\acknowledgementtext`、`\referenceentry...` 和 `\appendix...` 宏。

导出后建议在 Word 中做一次域更新：

- 全选全文后按 `F9`
- 或右键目录选择“更新域”

## 常改字段

打开 `main.tex` 顶部，优先修改这些宏：

- `\thesistitlecn`
- `\thesissubtitlecn`
- `\thesistitleen`
- `\thesissubtitleen`
- `\studentname`
- `\studentid`
- `\college`
- `\grade`
- `\major`
- `\company`
- `\advisor`
- `\cnabstracttext`
- `\cnkeywordslineone`
- `\cnkeywordslinetwo`
- `\enabstracttext`
- `\enkeywordslineone`
- `\enkeywordslinetwo`
- `\conclusiontext`
- `\acknowledgementtext`
- `\referencetitle`
- `\referenceentryone` 到 `\referenceentryten`
- `\appendixtitle`
- `\appendixnote`
- `\appendixatitle`
- `\appendixbtitle`

## 备选 Pandoc 模式

如果你想改用 `pandoc` 路线：

```bash
cd ..
bash scripts/install_pandoc.sh
python3 scripts/export_word.py --mode pandoc
```

`scripts/install_pandoc.sh` 在下载时会自动关闭代理环境变量。

## 说明

- 模板默认按 A4、正文宋体小四、页眉横线、页脚居中页码进行配置。
- 图表标题、目录标题、章节层级已经尽量向原 Word 模板靠齐。
- 当前仓库已经完成 LaTeX 编译校验和 Word 导出校验。
- 如果用于公开仓库，建议不要直接公开分发字体文件。
