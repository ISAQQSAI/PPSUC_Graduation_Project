# PPSUC Graduation Project LaTeX Template

中国人民公安大学信息网络安全学院本科毕业论文（设计）LaTeX 模板项目。

这个仓库以学校官方提供的 [信网学院本科毕业设计模板（2024）.docx](信网学院本科毕业设计模板（2024）.docx) 和 [信网学院本科毕业设计模板（2024）.pdf](信网学院本科毕业设计模板（2024）.pdf) 为参考，整理出一套可编译、可版本管理、可导出 Word 的模板工程。

## 项目概览

- 参考学校官方 Word / PDF 模板进行版式对齐
- 提供可直接编译的 XeLaTeX 工程
- 保留宋体、黑体、楷体、仿宋等中文字体方案
- 支持从同一个 `main.tex` 自动导出 `PDF` 和 `DOCX`
- 最终导出的 `PDF` 和 `DOCX` 默认输出到仓库根目录
- `DOCX` 导出默认复用官方 Word 模板结构，版式更接近学校模板
- 正文公式会导出为原生 Word 可编辑公式（OMML）

## 当前状态

目前这套仓库已经可以稳定完成两条输出链路：

1. `LaTeX -> PDF`
2. `LaTeX -> DOCX`

其中：

- `PDF` 模板采用“官方固定样张页 + LaTeX 可编辑正文页”的混合方案
- `DOCX` 模板采用“官方 Word 模板 + 从 `main.tex` 自动同步内容”的方案

这意味着现在的推荐工作方式已经比较明确：

1. 只维护 [latex-template/main.tex](latex-template/main.tex)
2. 用它编译出最终 PDF
3. 用它导出可继续编辑的 Word 版本

## 仓库结构

```text
PPSUC_Graduation_Project/
├── README.md
├── main.pdf
├── main.docx
├── docs/
│   └── template-comparison.png
├── latex-template/
│   ├── README.md
│   ├── main.tex
│   ├── assets/
│   │   ├── official-template.pdf
│   │   ├── pup-title.png
│   │   ├── pup-emblem-gray.png
│   │   └── sample-framework.png
│   └── fonts/
│       ├── SIMSUN.TTC
│       ├── SIMHEI.TTF
│       ├── 楷体_GB2312.TTF
│       └── 仿宋_GB2312.TTF
├── scripts/
│   ├── export_pdf.py
│   ├── export_word.py
│   └── install_pandoc.sh
├── tools/
│   └── pandoc
├── 信网学院本科毕业设计模板（2024）.docx
└── 信网学院本科毕业设计模板（2024）.pdf
```

## 快速开始

### 1. 编译 PDF

```bash
python3 scripts/export_pdf.py
```

输出文件：

```text
main.pdf
```

如果你只想在模板目录里手动编译调试：

```bash
cd latex-template
xelatex main.tex
xelatex main.tex
```

### 2. 导出 Word

在仓库根目录执行：

```bash
python3 scripts/export_word.py
```

默认输出：

```text
main.docx
```

如果你想自定义输出路径：

```bash
python3 scripts/export_word.py /path/to/output.docx
```

## Word 同步范围

默认 `template` 模式下，Word 导出会从 [latex-template/main.tex](latex-template/main.tex) 自动同步这些内容：

- 封面基础字段
- 中文摘要、英文摘要
- 正文章节标题
- 正文段落
- 原生 Word 公式（OMML）
- 表格标题和表格内容
- 图片与图题
- 代码块内容
- 结论、致谢
- 参考文献
- 附录标题和附录说明

默认模式的特点：

- 直接复用官方 `docx` 模板结构
- 版式对齐程度高于普通 Pandoc 导出
- 更适合作为“贴近学校模板的 Word 初稿”

## 推荐工作流

推荐把 [latex-template/main.tex](latex-template/main.tex) 当成唯一内容源：

1. 修改 `main.tex` 顶部论文信息和正文内容
2. 运行 `python3 scripts/export_pdf.py` 生成 PDF
3. 运行 `python3 scripts/export_word.py` 生成 DOCX
4. 在 Word 中更新目录和页码域
5. 做最后一轮人工检查后提交

Word 中建议再做一次域更新：

- 全选全文后按 `F9`
- 或右键目录选择“更新域”

这样目录标题、页码和交叉引用会更稳。

## 你最常需要修改的内容

打开 [latex-template/main.tex](latex-template/main.tex)，优先修改这些宏：

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
- `\cnabstracttitle`
- `\cnabstractsubtitle`
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

## 环境要求

建议环境：

- `xelatex`
- `latexmk`
- `ctex` 相关中文宏包
- Python 3

可选环境：

- `pandoc`
  说明：仅在使用 `--mode pandoc` 备选导出时需要

本项目依赖本地字体文件：

- `SIMSUN.TTC`
- `SIMHEI.TTF`
- `楷体_GB2312.TTF`
- `仿宋_GB2312.TTF`

字体默认放在 `latex-template/fonts/` 下，模板会直接从该目录加载。

## 已知说明

- 当前 PDF 仍是“官方固定页 + LaTeX 正文页”的混合方案，不是整份 PDF 都由 LaTeX 从零重建。
- 当前 DOCX 已经能从 `main.tex` 同步大部分核心内容，但导出后仍建议在 Word 中做一轮人工检查。
- 代码块、复杂自由排版内容、以及 Word 自己的目录页码，仍建议以最终导出文件为准做复核。

## 模板对比

下图左侧为导出的 Word 版，中间为官方 PDF 模板，右侧为当前 LaTeX 编译输出。

![Official PDF vs compiled LaTeX output](docs/template-comparison.png)
