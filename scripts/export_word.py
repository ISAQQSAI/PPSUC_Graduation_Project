#!/usr/bin/env python3
from __future__ import annotations

import argparse
import copy
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

try:
    from PIL import Image
except ImportError:  # pragma: no cover - optional dependency
    Image = None


ROOT = Path(__file__).resolve().parents[1]
LATEX_DIR = ROOT / "latex-template"
MAIN_TEX = LATEX_DIR / "main.tex"
REFERENCE_DOCX = ROOT / "信网学院本科毕业设计模板（2024）.docx"
DEFAULT_OUTPUT = ROOT / "main.docx"
LOCAL_PANDOC = ROOT / "tools" / "pandoc"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

MACRO_NAMES = [
    "thesistitlecn",
    "thesissubtitlecn",
    "cnabstracttitle",
    "cnabstractsubtitle",
    "thesistitleen",
    "thesissubtitleen",
    "studentname",
    "studentid",
    "college",
    "grade",
    "major",
    "company",
    "advisor",
    "cnabstracttext",
    "cnkeywordslineone",
    "cnkeywordslinetwo",
    "enabstracttext",
    "enkeywordslineone",
    "enkeywordslinetwo",
    "bodyplaceholder",
    "conclusiontext",
    "acknowledgementtext",
    "referencetitle",
    "referenceentryone",
    "referenceentrytwo",
    "referenceentrythree",
    "referenceentryfour",
    "referenceentryfive",
    "referenceentrysix",
    "referenceentryseven",
    "referenceentryeight",
    "referenceentrynine",
    "referenceentryten",
    "appendixtitle",
    "appendixnote",
    "appendixatitle",
    "appendixbtitle",
]

DIRECT_TEMPLATE_REPLACEMENTS = {
    "［键入论文题目］": "thesistitlecn",
    "［键入论文副标题或删除此行］": "thesissubtitlecn",
    "［键入学生姓名］": "studentname",
    "［键入学号］": "studentid",
    "［键入学院名称］": "college",
    "［键入年级］": "grade",
    "［键入专业名称］": "major",
    "［键入区队］": "company",
    "［键入指导教师姓名］": "advisor",
    "［单击此处键入标题］": "cnabstracttitle",
    "［单击此处继续键入副标题或删除此行］": "cnabstractsubtitle",
    "［中文摘要内容一般为200-300字左右，五号，仿宋］": "cnabstracttext",
    "［单击此处键入英文标题］": "thesistitleen",
    "［单击此处继续键入英文副标题或删除此行］": "thesissubtitleen",
    "［单击此处键入英文摘要内容，与中文摘要相对应，一般不宜超过250个实词］": "enabstracttext",
}

KEYWORD_PLACEHOLDER_PATTERN = re.compile(
    "|".join(
        re.escape(item)
        for item in [
            "［五号，仿宋］",
            "［单击此处键入关键词］",
            "［单击此处键入关键词或删除］",
        ]
    )
)


def load_macros(tex_path: Path) -> dict[str, str]:
    text = tex_path.read_text(encoding="utf-8")
    return load_macros_from_text(text)


def load_macros_from_text(text: str) -> dict[str, str]:
    macros: dict[str, str] = {}
    for name in MACRO_NAMES:
        pattern = re.compile(r"\\newcommand\{\\%s\}\{((?:[^{}]|\{[^{}]*\})*)\}" % re.escape(name))
        match = pattern.search(text)
        if not match:
            macros[name] = ""
            continue
        macros[name] = clean_tex_text(match.group(1))
    return macros


def clean_tex_text(value: str) -> str:
    replacements = {
        r"\quad": "　",
        r"\par": "\n",
        r"\\": "\n",
        r"\_": "_",
        r"\&": "&",
        r"\%": "%",
        r"\#": "#",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    value = re.sub(r"\\[a-zA-Z]+\*?(?:\{[^{}]*\})?", "", value)
    value = value.replace("{", "").replace("}", "")
    value = re.sub(r"[ \t]+", " ", value)
    value = value.replace("　 ", "　").replace(" 　", "　")
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def safe_xml_text(value: str) -> str:
    return escape(value.replace("\n", " ").strip())


def expand_tex_macros(value: str, macros: dict[str, str]) -> str:
    for name, macro_value in macros.items():
        value = value.replace(f"\\{name}", macro_value)
    return value


def parse_keywords(*values: str) -> list[str]:
    items: list[str] = []
    for value in values:
        for token in re.split(r"[；;]", clean_tex_text(value)):
            token = token.strip()
            if not token:
                continue
            if token.startswith("［") and token.endswith("］"):
                continue
            items.append(token)
    return items


def is_working_pandoc(path: str) -> bool:
    try:
        subprocess.run(
            [path, "--version"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except (OSError, subprocess.CalledProcessError):
        return False
    return True


def choose_pandoc() -> str | None:
    if LOCAL_PANDOC.exists() and is_working_pandoc(str(LOCAL_PANDOC)):
        return str(LOCAL_PANDOC)

    system_pandoc = shutil.which("pandoc")
    if system_pandoc and is_working_pandoc(system_pandoc):
        return system_pandoc

    return None


def build_markdown(macros: dict[str, str]) -> str:
    cn_keywords = "；".join(parse_keywords(macros["cnkeywordslineone"], macros["cnkeywordslinetwo"]))
    en_keywords = "；".join(parse_keywords(macros["enkeywordslineone"], macros["enkeywordslinetwo"]))

    return f"""---
title: "{escape_yaml(macros['thesistitlecn'])}"
subtitle: "{escape_yaml(macros['thesissubtitlecn'])}"
author: "{escape_yaml(macros['studentname'])}"
lang: zh-CN
toc: true
toc-depth: 2
---

# 论文信息

| 字段 | 内容 |
| --- | --- |
| 中文题目 | {macros['thesistitlecn']} |
| 中文副标题 | {macros['thesissubtitlecn']} |
| 英文题目 | {macros['thesistitleen']} |
| 英文副标题 | {macros['thesissubtitleen']} |
| 学生姓名 | {macros['studentname']} |
| 学号 | {macros['studentid']} |
| 学院 | {macros['college']} |
| 年级 | {macros['grade']} |
| 专业 | {macros['major']} |
| 区队 | {macros['company']} |
| 指导教师 | {macros['advisor']} |

# 中文摘要

## 标题

{macros['cnabstracttitle']}

## 副标题

{macros['cnabstractsubtitle']}

## 摘要正文

{macros['cnabstracttext']}

## 关键词

{cn_keywords}

# English Abstract

## Title

{macros['thesistitleen']}

## Subtitle

{macros['thesissubtitleen']}

## Abstract

{macros['enabstracttext']}

## Keywords

{en_keywords}

# 正文

## 1. 引言

［单击此处键入内容］

## 2. 【单击此处键入一级标题】

### 2.1 【单击此处键入二级标题】

#### 2.1.1 ［单击此处键入内容］

2.1.1.1 ［单击此处键入内容］

### 2.2 【单击此处键入二级标题】

［单击此处键入内容］

## 3. 【单击此处键入一级标题】

### 3.1 【单击此处键入二级标题】

［单击此处键入内容］

## 4. 结论

［单击此处键入内容］

# 致谢

［单击此处键入致谢词］

# 参考文献

[1] [序号] 主要责任者. 题名[文献类型标识]. 其他责任者. 版本项. 出版地: 出版者, 出版年.

[2] [序号] 析出文献主要责任者. 析出文献题名[文献类型标识]. 期刊题名, 年, 卷(期): 页码.

# 附录

## 附录 A 图目录

图目录内容可按需要补充。

## 附录 B 表目录

表目录内容可按需要补充。
"""


def escape_yaml(value: str) -> str:
    return value.replace('"', '\\"')


def replace_keyword_placeholders(text: str, keywords: list[str]) -> str:
    if not keywords:
        return text

    items = iter([safe_xml_text(item) for item in keywords])

    def repl(match: re.Match[str]) -> str:
        try:
            return next(items)
        except StopIteration:
            return match.group(0)

    return KEYWORD_PLACEHOLDER_PATTERN.sub(repl, text)


def extract_command_arguments(text: str, command: str) -> list[str]:
    pattern = re.compile(r"\\%s\{((?:[^{}]|\{[^{}]*\})*)\}" % re.escape(command))
    return [match.group(1) for match in pattern.finditer(text)]


def strip_prefix(text: str, prefix: str) -> str:
    if text.startswith(prefix):
        text = text[len(prefix):]
    return text.strip()


def replace_sequential_occurrences(text: str, placeholder: str, replacements: list[str]) -> str:
    start = 0
    for replacement in replacements:
        index = text.find(placeholder, start)
        if index == -1:
            break
        text = text[:index] + replacement + text[index + len(placeholder) :]
        start = index + len(replacement)
    return text


def word_tag(name: str) -> str:
    return f"{{{WORD_NS}}}{name}"


def math_tag(name: str) -> str:
    return f"{{{MATH_NS}}}{name}"


def word_attr(name: str) -> str:
    return f"{{{WORD_NS}}}{name}"


def get_or_create_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.SubElement(parent, tag)
    return child


def remove_child(parent: ET.Element, tag: str) -> None:
    child = parent.find(tag)
    if child is not None:
        parent.remove(child)


def paragraph_props(paragraph: ET.Element) -> ET.Element:
    return get_or_create_child(paragraph, word_tag("pPr"))


def paragraph_run_props(paragraph: ET.Element) -> ET.Element:
    return get_or_create_child(paragraph_props(paragraph), word_tag("rPr"))


def set_word_toggle(rpr: ET.Element, name: str, enabled: bool | None) -> None:
    if enabled is None:
        return
    tag = word_tag(name)
    remove_child(rpr, tag)
    node = ET.SubElement(rpr, tag)
    if not enabled:
        node.set(word_attr("val"), "0")


def set_word_run_format(
    rpr: ET.Element,
    *,
    ascii_font: str | None = None,
    hansi_font: str | None = None,
    east_asia_font: str | None = None,
    cs_font: str | None = None,
    size: int | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    if any(value is not None for value in (ascii_font, hansi_font, east_asia_font, cs_font)):
        rfonts = get_or_create_child(rpr, word_tag("rFonts"))
        if ascii_font is not None:
            rfonts.set(word_attr("ascii"), ascii_font)
        if hansi_font is not None:
            rfonts.set(word_attr("hAnsi"), hansi_font)
        if east_asia_font is not None:
            rfonts.set(word_attr("eastAsia"), east_asia_font)
            rfonts.set(word_attr("hint"), "eastAsia")
        if cs_font is not None:
            rfonts.set(word_attr("cs"), cs_font)

    if size is not None:
        get_or_create_child(rpr, word_tag("sz")).set(word_attr("val"), str(size))
        get_or_create_child(rpr, word_tag("szCs")).set(word_attr("val"), str(size))

    set_word_toggle(rpr, "b", bold)
    set_word_toggle(rpr, "bCs", bold)
    set_word_toggle(rpr, "i", italic)
    set_word_toggle(rpr, "iCs", italic)


def set_paragraph_alignment(paragraph: ET.Element, value: str) -> None:
    get_or_create_child(paragraph_props(paragraph), word_tag("jc")).set(word_attr("val"), value)


def set_page_break_before(paragraph: ET.Element, enabled: bool) -> None:
    ppr = paragraph_props(paragraph)
    tag = word_tag("pageBreakBefore")
    remove_child(ppr, tag)
    if enabled:
        ET.SubElement(ppr, tag)


def set_paragraph_spacing(
    paragraph: ET.Element,
    *,
    line: int | None = None,
    line_rule: str | None = None,
    before: int | None = None,
    after: int | None = None,
) -> None:
    spacing = get_or_create_child(paragraph_props(paragraph), word_tag("spacing"))
    if line is not None:
        spacing.set(word_attr("line"), str(line))
    if line_rule is not None:
        spacing.set(word_attr("lineRule"), line_rule)
    if before is not None:
        spacing.set(word_attr("before"), str(before))
    if after is not None:
        spacing.set(word_attr("after"), str(after))


def set_paragraph_indent(
    paragraph: ET.Element,
    *,
    left: int | None = None,
    left_chars: int | None = None,
    first_line: int | None = None,
    first_line_chars: int | None = None,
) -> None:
    ind = get_or_create_child(paragraph_props(paragraph), word_tag("ind"))
    if left is not None:
        ind.set(word_attr("left"), str(left))
    if left_chars is not None:
        ind.set(word_attr("leftChars"), str(left_chars))
    if first_line is not None:
        ind.set(word_attr("firstLine"), str(first_line))
    if first_line_chars is not None:
        ind.set(word_attr("firstLineChars"), str(first_line_chars))


def apply_paragraph_text_format(
    paragraph: ET.Element,
    *,
    ascii_font: str,
    hansi_font: str,
    east_asia_font: str,
    cs_font: str,
    size: int,
    bold: bool | None = None,
    italic: bool | None = None,
) -> None:
    set_word_run_format(
        paragraph_run_props(paragraph),
        ascii_font=ascii_font,
        hansi_font=hansi_font,
        east_asia_font=east_asia_font,
        cs_font=cs_font,
        size=size,
        bold=bold,
        italic=italic,
    )

    for run in paragraph.iter(word_tag("r")):
        set_word_run_format(
            get_or_create_child(run, word_tag("rPr")),
            ascii_font=ascii_font,
            hansi_font=hansi_font,
            east_asia_font=east_asia_font,
            cs_font=cs_font,
            size=size,
            bold=bold,
            italic=italic,
        )


def apply_body_paragraph_format(paragraph: ET.Element) -> None:
    ppr = paragraph_props(paragraph)
    for name in ("pStyle", "outlineLvl", "keepNext", "keepLines"):
        remove_child(ppr, word_tag(name))
    set_paragraph_spacing(paragraph, line=360, line_rule="auto", before=0, after=0)
    set_paragraph_indent(paragraph, left=0, left_chars=0, first_line=460, first_line_chars=200)
    set_paragraph_alignment(paragraph, "both")
    apply_paragraph_text_format(
        paragraph,
        ascii_font="Times New Roman",
        hansi_font="Times New Roman",
        east_asia_font="宋体",
        cs_font="Times New Roman",
        size=24,
        bold=False,
        italic=False,
    )


def apply_caption_paragraph_format(paragraph: ET.Element) -> None:
    ppr = paragraph_props(paragraph)
    remove_child(ppr, word_tag("pStyle"))
    remove_child(ppr, word_tag("outlineLvl"))
    set_paragraph_spacing(paragraph, line=300, line_rule="auto", before=0, after=0)
    set_paragraph_indent(paragraph, left=0, left_chars=0, first_line=0, first_line_chars=0)
    set_paragraph_alignment(paragraph, "center")
    apply_paragraph_text_format(
        paragraph,
        ascii_font="Times New Roman",
        hansi_font="Times New Roman",
        east_asia_font="宋体",
        cs_font="Times New Roman",
        size=21,
        bold=False,
        italic=False,
    )


def apply_formula_paragraph_format(paragraph: ET.Element) -> None:
    ppr = paragraph_props(paragraph)
    remove_child(ppr, word_tag("pStyle"))
    set_paragraph_indent(paragraph, left=0, left_chars=0, first_line=0, first_line_chars=0)
    set_paragraph_alignment(paragraph, "center")
    apply_paragraph_text_format(
        paragraph,
        ascii_font="Times New Roman",
        hansi_font="Times New Roman",
        east_asia_font="宋体",
        cs_font="Times New Roman",
        size=24,
        bold=False,
        italic=False,
    )


def apply_table_format(table: ET.Element) -> None:
    for cell in table.iter(word_tag("tc")):
        for paragraph in cell.findall(word_tag("p")):
            set_paragraph_spacing(paragraph, line=300, line_rule="auto", before=0, after=0)
            set_paragraph_indent(paragraph, left=0, left_chars=0, first_line=0, first_line_chars=0)
            set_paragraph_alignment(paragraph, "center")
            apply_paragraph_text_format(
                paragraph,
                ascii_font="Times New Roman",
                hansi_font="Times New Roman",
                east_asia_font="宋体",
                cs_font="Times New Roman",
                size=21,
                bold=False,
                italic=False,
            )


def apply_code_paragraph_format(paragraph: ET.Element) -> None:
    apply_paragraph_text_format(
        paragraph,
        ascii_font="Consolas",
        hansi_font="Consolas",
        east_asia_font="宋体",
        cs_font="Consolas",
        size=18,
        bold=False,
        italic=False,
    )


def apply_format_to_top_level_paragraphs(
    body: ET.Element,
    start: int,
    end: int,
    formatter,
) -> None:
    children = list(body)
    for idx in range(start, min(end, len(children))):
        if children[idx].tag == word_tag("p"):
            formatter(children[idx])


def numbered_title_suffix(title: str) -> str:
    match = re.match(r"^\d+(?:\.\d+)?\s*(.*)$", title)
    if match:
        return match.group(1).strip()
    return title.strip()


def replace_split_body_title(document_xml: str, target_full_title: str, replacement_full_title: str) -> str:
    root = ET.fromstring(document_xml)

    for paragraph in root.iter(word_tag("p")):
        text_nodes = [node for node in paragraph.iter() if node.tag == word_tag("t")]
        joined = "".join(node.text or "" for node in text_nodes).strip()
        if joined != target_full_title:
            continue

        replacement_suffix = numbered_title_suffix(replacement_full_title)
        if len(text_nodes) >= 2:
            text_nodes[1].text = replacement_suffix
            for node in text_nodes[2:]:
                node.text = ""
        elif text_nodes:
            text_nodes[0].text = replacement_full_title
        break

    return ET.tostring(root, encoding="unicode")


def paragraph_text_nodes(paragraph: ET.Element) -> list[ET.Element]:
    nodes: list[ET.Element] = []

    def walk(element: ET.Element) -> None:
        for child in element:
            if child.tag == word_tag("p"):
                continue
            if child.tag == word_tag("t"):
                nodes.append(child)
                continue
            walk(child)

    walk(paragraph)
    return nodes


def paragraph_has_section_break(paragraph: ET.Element) -> bool:
    if paragraph.tag != word_tag("p"):
        return False
    ppr = paragraph.find(word_tag("pPr"))
    if ppr is None:
        return False
    return ppr.find(word_tag("sectPr")) is not None


def split_preserved_paragraph_children(paragraph: ET.Element) -> tuple[list[ET.Element], list[ET.Element]]:
    preserved_tags = {
        word_tag("bookmarkStart"),
        word_tag("bookmarkEnd"),
        word_tag("commentRangeStart"),
        word_tag("commentRangeEnd"),
        word_tag("proofErr"),
        word_tag("permStart"),
        word_tag("permEnd"),
    }
    paragraph_props = paragraph.find(word_tag("pPr"))
    children = [child for child in list(paragraph) if child is not paragraph_props]

    prefix: list[ET.Element] = []
    index = 0
    while index < len(children) and children[index].tag in preserved_tags:
        prefix.append(copy.deepcopy(children[index]))
        index += 1

    suffix: list[ET.Element] = []
    end_index = len(children) - 1
    while end_index >= index and children[end_index].tag in preserved_tags:
        suffix.append(copy.deepcopy(children[end_index]))
        end_index -= 1
    suffix.reverse()

    return prefix, suffix


def paragraph_visible_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph_text_nodes(paragraph)).strip()


def set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = paragraph_text_nodes(paragraph)
    if not text_nodes:
        return
    text_nodes[0].text = text
    for node in text_nodes[1:]:
        node.text = ""


def apply_reference_appendix_sync(document_xml: str, sync_data: dict[str, object]) -> str:
    root = ET.fromstring(document_xml)
    paragraphs = [paragraph for paragraph in root.iter(word_tag("p"))]
    visible_texts = [paragraph_visible_text(paragraph) for paragraph in paragraphs]

    reference_guidance_texts = {
        "专著（包括以各种载体形式出版的普通图书、学位论文、技术报告、会议文集、汇编）不用此信息时，删除此框。",
        "专著中的析出文献具体格式，不用此信息时，删除此框。",
        "期刊中析出的文献，不用此信息时，删除此框。",
        "报纸文章参考文献具体格式，不用此信息时，删除此框。",
        "专利文献，不用此信息时，删除此框。",
        "电子资源著录具体格式，不用此信息时，删除此框。",
        "各种未定类型参考文献具体格式，不用此信息时，删除此框。",
    }
    reference_markers = [
        "[1]",
        "[2]",
        "[3]",
        "[4]",
        "[5]",
        "[6]",
        "（1）[序号]",
        "（2）[序号]",
        "（3）[序号]",
        "[7]",
    ]

    reference_title = sync_data["reference_title"]
    reference_entries = sync_data["reference_entries"]
    appendix_title = sync_data["appendix_title"]
    appendix_note = sync_data["appendix_note"]
    appendix_a_title = sync_data["appendix_a_title"]
    appendix_b_title = sync_data["appendix_b_title"]

    reference_start = next(
        (idx for idx, text in enumerate(visible_texts) if "参考文献：" in text),
        None,
    )
    appendix_start = next(
        (idx for idx, text in enumerate(visible_texts) if "附  录：" in text),
        None,
    )

    if reference_start is not None:
        set_paragraph_text(paragraphs[reference_start], reference_title)
        entry_index = 0
        stop_index = appendix_start if appendix_start is not None else len(paragraphs)
        for idx in range(reference_start + 1, stop_index):
            text = visible_texts[idx]
            if not text:
                continue
            if text in reference_guidance_texts:
                set_paragraph_text(paragraphs[idx], "")
                continue
            if entry_index < len(reference_markers) and reference_markers[entry_index] in text:
                replacement = reference_entries[entry_index] if entry_index < len(reference_entries) else ""
                set_paragraph_text(paragraphs[idx], replacement)
                entry_index += 1

    if appendix_start is not None:
        set_paragraph_text(paragraphs[appendix_start], appendix_title)
        note_replaced = False
        appendix_a_replaced = False
        appendix_b_replaced = False
        for idx in range(appendix_start + 1, len(paragraphs)):
            text = visible_texts[idx]
            if not text:
                continue
            if text in {
                "电子资源著录具体格式，不用此信息时，删除此框。",
                "各种未定类型参考文献具体格式，不用此信息时，删除此框。",
            }:
                set_paragraph_text(paragraphs[idx], "")
                continue
            if not note_replaced and text == "（信网统一格式，供参考）":
                set_paragraph_text(paragraphs[idx], appendix_note)
                note_replaced = True
                continue
            if not appendix_a_replaced and "附录A" in text and "图目录" in text:
                set_paragraph_text(paragraphs[idx], appendix_a_title)
                appendix_a_replaced = True
                continue
            if not appendix_b_replaced and "附录" in text and "表目录" in text:
                set_paragraph_text(paragraphs[idx], appendix_b_title)
                appendix_b_replaced = True
                continue

    return ET.tostring(root, encoding="unicode")


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", "", text)


def top_level_paragraph_text(element: ET.Element) -> str:
    if element.tag != word_tag("p"):
        return ""
    return paragraph_visible_text(element)


def find_top_level_index(
    body: ET.Element,
    target_text: str,
    *,
    start: int = 0,
    contains: bool = False,
) -> int:
    target = normalize_spaces(target_text)
    children = list(body)
    for idx in range(start, len(children)):
        text = top_level_paragraph_text(children[idx])
        if not text:
            continue
        normalized = normalize_spaces(text)
        if contains:
            if target in normalized:
                return idx
        elif normalized == target:
            return idx
    raise ValueError(f"Unable to locate paragraph: {target_text}")


def clone_first_run_properties(paragraph: ET.Element) -> ET.Element | None:
    for run in paragraph.iter(word_tag("r")):
        run_props = run.find(word_tag("rPr"))
        if run_props is not None:
            return copy.deepcopy(run_props)
    return None


def reset_paragraph_text(
    paragraph: ET.Element,
    text: str,
    *,
    source_paragraph: ET.Element | None = None,
) -> None:
    paragraph_props = paragraph.find(word_tag("pPr"))
    preserved_prefix, preserved_suffix = split_preserved_paragraph_children(paragraph)
    for child in list(paragraph):
        if child is paragraph_props:
            continue
        paragraph.remove(child)

    for child in preserved_prefix:
        paragraph.append(child)

    if not text:
        for child in preserved_suffix:
            paragraph.append(child)
        return

    run = ET.SubElement(paragraph, word_tag("r"))
    run_props = clone_first_run_properties(source_paragraph if source_paragraph is not None else paragraph)
    if run_props is not None:
        run.append(run_props)

    text_node = ET.SubElement(run, word_tag("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(XML_SPACE, "preserve")
    text_node.text = text

    for child in preserved_suffix:
        paragraph.append(child)


def make_paragraph_like(template: ET.Element, text: str) -> ET.Element:
    paragraph = copy.deepcopy(template)
    reset_paragraph_text(paragraph, text, source_paragraph=template)
    return paragraph


def replace_top_level_range(body: ET.Element, start: int, end: int, new_elements: list[ET.Element]) -> None:
    children = list(body)
    preserved_section_breaks = [
        copy.deepcopy(child) for child in children[start:end] if paragraph_has_section_break(child)
    ]
    for child in children[start:end]:
        body.remove(child)
    final_elements = list(new_elements) + preserved_section_breaks
    for offset, element in enumerate(final_elements):
        body.insert(start + offset, element)


def replace_paragraph_range_with_text(
    body: ET.Element,
    start: int,
    end: int,
    template: ET.Element,
    texts: list[str],
    *,
    trailing_blank_count: int = 0,
    blank_template: ET.Element | None = None,
) -> None:
    new_elements = [make_paragraph_like(template, text) for text in texts]
    if blank_template is not None:
        new_elements.extend(make_paragraph_like(blank_template, "") for _ in range(trailing_blank_count))
    replace_top_level_range(body, start, end, new_elements)


def replace_paragraph_text_nodes(text_nodes: list[ET.Element], text: str) -> None:
    if not text_nodes:
        return
    text_nodes[0].text = text
    if text.startswith(" ") or text.endswith(" "):
        text_nodes[0].set(XML_SPACE, "preserve")
    else:
        text_nodes[0].attrib.pop(XML_SPACE, None)
    for node in text_nodes[1:]:
        node.text = ""
        node.attrib.pop(XML_SPACE, None)


def next_top_level_paragraph_index(body: ET.Element, start: int, stop: int | None = None) -> int:
    children = list(body)
    limit = len(children) if stop is None else min(stop, len(children))
    for idx in range(start, limit):
        if children[idx].tag == word_tag("p"):
            return idx
    raise ValueError("Unable to locate next paragraph node")


def set_cell_text(cell: ET.Element, text: str) -> None:
    paragraphs = cell.findall(word_tag("p"))
    if not paragraphs:
        paragraphs = [ET.SubElement(cell, word_tag("p"))]
    reset_paragraph_text(paragraphs[0], text, source_paragraph=paragraphs[0])
    for paragraph in paragraphs[1:]:
        reset_paragraph_text(paragraph, "", source_paragraph=paragraph)


def update_table_rows(table: ET.Element, rows: list[list[str]]) -> None:
    if not rows:
        return

    table_rows = table.findall(word_tag("tr"))
    if not table_rows:
        return

    data_template = table_rows[1] if len(table_rows) > 1 else table_rows[0]

    while len(table_rows) < len(rows):
        table.append(copy.deepcopy(data_template))
        table_rows = table.findall(word_tag("tr"))

    while len(table_rows) > len(rows):
        table.remove(table_rows[-1])
        table_rows = table.findall(word_tag("tr"))

    for row_idx, row_values in enumerate(rows):
        cells = table_rows[row_idx].findall(word_tag("tc"))
        for cell_idx, value in enumerate(row_values):
            if cell_idx < len(cells):
                set_cell_text(cells[cell_idx], value)
        for cell in cells[len(row_values) :]:
            set_cell_text(cell, "")


def extract_environment_bodies(text: str, env_name: str) -> list[str]:
    start_token = f"\\begin{{{env_name}}}"
    end_token = f"\\end{{{env_name}}}"
    bodies: list[str] = []
    cursor = 0

    while True:
        start = text.find(start_token, cursor)
        if start == -1:
            break
        body_start = start + len(start_token)

        if env_name == "tabular" and body_start < len(text) and text[body_start] == "{":
            depth = 0
            pos = body_start
            while pos < len(text):
                char = text[pos]
                if char == "{":
                    depth += 1
                elif char == "}":
                    depth -= 1
                    if depth == 0:
                        pos += 1
                        break
                pos += 1
            body_start = pos

        end = text.find(end_token, body_start)
        if end == -1:
            break

        bodies.append(text[body_start:end])
        cursor = end + len(end_token)

    return bodies


def equation_block_to_display_math(block: str) -> str:
    tag_match = re.search(r"\\tag\{([^}]*)\}", block)
    equation_number = tag_match.group(1).strip() if tag_match else ""
    block_without_tag = re.sub(r"\\tag\{[^}]*\}", "", block).strip()
    if equation_number:
        block_without_tag = f"{block_without_tag}\\qquad({equation_number})"
    return f"\\[\n{block_without_tag}\n\\]"


def extract_equation_number(block: str) -> str:
    tag_match = re.search(r"\\tag\{([^}]*)\}", block)
    return tag_match.group(1).strip() if tag_match else ""


def convert_equation_block_to_omml_xml(block: str) -> str | None:
    pandoc = choose_pandoc()
    if pandoc is None:
        return None

    latex_source = "\n".join(
        [
            r"\documentclass{article}",
            r"\usepackage{amsmath}",
            r"\begin{document}",
            equation_block_to_display_math(block),
            r"\end{document}",
            "",
        ]
    )

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            tex_path = tmpdir_path / "equation.tex"
            docx_path = tmpdir_path / "equation.docx"
            tex_path.write_text(latex_source, encoding="utf-8")
            subprocess.run(
                [pandoc, "--from", "latex", str(tex_path), "-o", str(docx_path)],
                check=True,
                capture_output=True,
                text=True,
            )
            with zipfile.ZipFile(docx_path, "r") as archive:
                root = ET.fromstring(archive.read("word/document.xml"))
    except (OSError, subprocess.CalledProcessError, zipfile.BadZipFile, ET.ParseError):
        return None

    body = root.find(word_tag("body"))
    if body is None:
        return None

    for child in list(body):
        if child.find(f".//{math_tag('oMathPara')}") is not None or child.find(
            f".//{math_tag('oMath')}"
        ) is not None:
            return ET.tostring(child, encoding="unicode")
    return None


def parse_tabular_rows(body: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for raw_line in body.splitlines():
        line = raw_line.strip()
        if not line or line == r"\hline":
            continue
        if line.endswith(r"\\"):
            line = line[:-2].strip()
        if not line or line == r"\hline":
            continue
        cells = [clean_tex_text(cell).replace("\n", " ").strip() for cell in line.split("&")]
        rows.append(cells)
    return rows


def parse_equation_lines(block: str) -> list[str]:
    tag_match = re.search(r"\\tag\{([^}]*)\}", block)
    equation_number = f"({tag_match.group(1).strip()})" if tag_match else ""

    aligned_match = re.search(r"\\begin\{aligned\}(.*?)\\end\{aligned\}", block, re.S)
    inner = aligned_match.group(1) if aligned_match else block
    lines = []
    for raw_line in re.split(r"\\\\", inner):
        line = raw_line.strip()
        if not line:
            continue
        line = line.replace("&", "")
        line = line.replace(r"\quad", " ")
        line = line.replace(r"\,", " ")
        line = re.sub(r"\s+", " ", line)
        lines.append(line.strip())

    if equation_number:
        if lines:
            lines[-1] = f"{lines[-1]}    {equation_number}"
        else:
            lines.append(equation_number)

    return lines


def logical_paragraphs(text: str) -> list[str]:
    paragraphs = [re.sub(r"\s+", " ", item).strip() for item in re.split(r"\n+", text) if item.strip()]
    return paragraphs or [""]


def load_figure_image_bytes(path: Path) -> bytes | None:
    if not path.exists():
        return None
    if path.suffix.lower() == ".png":
        return path.read_bytes()
    if Image is None:
        return None

    image = Image.open(path)
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def make_normal_math_text_run(text: str) -> ET.Element:
    run = ET.Element(math_tag("r"))
    run_props = ET.SubElement(run, math_tag("rPr"))
    ET.SubElement(run_props, math_tag("nor"))
    text_node = ET.SubElement(run, math_tag("t"))
    text_node.text = text
    return run


def normalize_equation_number_in_omml(omml_paragraph: ET.Element, equation_number: str) -> None:
    if not equation_number:
        return

    o_math = omml_paragraph.find(f".//{math_tag('oMath')}")
    if o_math is None:
        return

    children = list(o_math)
    replacement = make_normal_math_text_run(f"({equation_number})")
    for idx in range(len(children) - 1, -1, -1):
        child = children[idx]
        if child.tag == math_tag("d"):
            o_math.remove(child)
            o_math.insert(idx, replacement)
            return

    o_math.append(replacement)


def make_omml_paragraph(
    omml_xml: str,
    template: ET.Element,
    equation_number: str = "",
) -> ET.Element | None:
    try:
        omml_paragraph = ET.fromstring(omml_xml)
    except ET.ParseError:
        return None

    normalize_equation_number_in_omml(omml_paragraph, equation_number)

    paragraph = copy.deepcopy(template)
    paragraph_props = paragraph.find(word_tag("pPr"))
    for child in list(paragraph):
        if child is paragraph_props:
            continue
        paragraph.remove(child)

    if paragraph_props is None:
        paragraph_props = ET.SubElement(paragraph, word_tag("pPr"))

    jc = paragraph_props.find(word_tag("jc"))
    if jc is None:
        jc = ET.SubElement(paragraph_props, word_tag("jc"))
    jc.set(word_attr("val"), "center")

    ind = paragraph_props.find(word_tag("ind"))
    if ind is not None:
        paragraph_props.remove(ind)

    for child in list(omml_paragraph):
        if child.tag == word_tag("pPr"):
            continue
        paragraph.append(copy.deepcopy(child))

    apply_formula_paragraph_format(paragraph)
    return paragraph


def build_equation_content_elements(
    sync_data: dict[str, object],
    intro_template: ET.Element,
    label_template: ET.Element,
    equation_template: ET.Element,
    blank_template: ET.Element,
) -> list[ET.Element]:
    elements = [make_paragraph_like(intro_template, sync_data["formula_intro"])]
    labels = sync_data["formula_example_labels"]
    equations = sync_data["equations"]
    equation_omml_xmls = sync_data["equation_omml_xmls"]
    equation_numbers = sync_data["equation_numbers"]
    count = max(len(labels), len(equations), len(equation_omml_xmls))

    for idx in range(count):
        label = labels[idx] if idx < len(labels) else f"示例 {idx + 1}:"
        elements.append(make_paragraph_like(label_template, label))

        omml_xml = equation_omml_xmls[idx] if idx < len(equation_omml_xmls) else None
        equation_number = equation_numbers[idx] if idx < len(equation_numbers) else ""
        if omml_xml:
            paragraph = make_omml_paragraph(omml_xml, equation_template, equation_number)
            if paragraph is not None:
                elements.append(paragraph)
            else:
                for line in equations[idx] if idx < len(equations) else []:
                    paragraph = make_paragraph_like(equation_template, line)
                    apply_formula_paragraph_format(paragraph)
                    elements.append(paragraph)
        else:
            for line in equations[idx] if idx < len(equations) else []:
                paragraph = make_paragraph_like(equation_template, line)
                apply_formula_paragraph_format(paragraph)
                elements.append(paragraph)

        elements.append(make_paragraph_like(blank_template, ""))

    return elements


def build_full_content_sync(tex_text: str, macros: dict[str, str]) -> dict[str, object]:
    document_text = tex_text.split(r"\begin{document}", 1)[-1]
    editable_body = document_text.split(r"\insertofficialpage{9}", 1)[0]

    manual_titles = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(editable_body, "manualtitle")
    ]
    manual_subsections = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(editable_body, "manualsubsection")
    ]
    manual_subsubsections = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(editable_body, "manualsubsubsection")
    ]
    manual_paragraphs = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(editable_body, "manualparagraph")
    ]

    heiti_pattern = re.compile(r"\{\\heiti\\zihao\{-4\}((?:[^{}]|\{[^{}]*\})*)\\par\}", re.S)
    heiti_paragraphs = [
        clean_tex_text(expand_tex_macros(match.group(1), macros))
        for match in heiti_pattern.finditer(editable_body)
    ]

    songti_pattern = re.compile(r"\{\\songti\\zihao\{-4\}((?:[^{}]|\{[^{}]*\})*)\\par\}", re.S)
    songti_paragraphs = [
        clean_tex_text(expand_tex_macros(match.group(1), macros))
        for match in songti_pattern.finditer(editable_body)
    ]

    caption_pattern = re.compile(r"\{\\songti\\zihao\{5\}([^{}]+?)\\par\}|\{\\songti\\zihao\{5\}([^{}]+?)\}", re.S)
    captions = [
        clean_tex_text(expand_tex_macros(match.group(1) or match.group(2), macros))
        for match in caption_pattern.finditer(editable_body)
    ]

    equation_blocks = extract_environment_bodies(editable_body, "equation")
    equations = [parse_equation_lines(block) for block in equation_blocks]
    equation_numbers = [extract_equation_number(block) for block in equation_blocks]
    equation_omml_xmls = [convert_equation_block_to_omml_xml(block) for block in equation_blocks]
    tables = [parse_tabular_rows(block) for block in extract_environment_bodies(editable_body, "tabular")]
    if len(tables) >= 3:
        second_table = tables[1] + tables[2]
        tables = [tables[0], second_table]

    image_match = re.search(r"\\includegraphics\[[^\]]*\]\{([^}]+)\}", editable_body)
    figure_image_path = None
    figure_image_bytes = None
    if image_match:
        figure_image_path = (LATEX_DIR / image_match.group(1)).resolve()
        figure_image_bytes = load_figure_image_bytes(figure_image_path)

    listing_match = re.search(r"\\begin\{lstlisting\}\n(.*?)\\end\{lstlisting\}", editable_body, re.S)
    code_lines = listing_match.group(1).rstrip("\n").splitlines() if listing_match else []

    return {
        "intro_title": manual_titles[0] if manual_titles else "1.引 言",
        "intro_paragraphs": logical_paragraphs(songti_paragraphs[0] if len(songti_paragraphs) > 0 else macros.get("bodyplaceholder", "")),
        "chapter_titles": [
            manual_titles[1] if len(manual_titles) > 1 else "2.标题",
            manual_titles[2] if len(manual_titles) > 2 else "3.标题",
        ],
        "section_titles": [
            manual_subsections[0] if len(manual_subsections) > 0 else "2.1 标题",
            manual_subsections[1] if len(manual_subsections) > 1 else "2.2 标题",
            manual_subsections[2] if len(manual_subsections) > 2 else "3.1 标题",
        ],
        "subsection_line": manual_subsubsections[0] if manual_subsubsections else "2.1.1",
        "paragraph_line": manual_paragraphs[0] if manual_paragraphs else "2.1.1.1",
        "section_body_paragraphs": [
            logical_paragraphs(songti_paragraphs[1] if len(songti_paragraphs) > 1 else macros.get("bodyplaceholder", "")),
            logical_paragraphs(songti_paragraphs[2] if len(songti_paragraphs) > 2 else macros.get("bodyplaceholder", "")),
        ],
        "formula_heading": heiti_paragraphs[0] if len(heiti_paragraphs) > 0 else "公式使用示例",
        "formula_intro": songti_paragraphs[3] if len(songti_paragraphs) > 3 else "",
        "formula_example_labels": [
            songti_paragraphs[4] if len(songti_paragraphs) > 4 else "示例 1:",
            songti_paragraphs[5] if len(songti_paragraphs) > 5 else "示例 2:",
        ],
        "equations": equations,
        "equation_numbers": equation_numbers,
        "equation_omml_xmls": equation_omml_xmls,
        "table_heading": heiti_paragraphs[1] if len(heiti_paragraphs) > 1 else "表格使用示例",
        "table_intro": songti_paragraphs[6] if len(songti_paragraphs) > 6 else "",
        "table_example_label": songti_paragraphs[7] if len(songti_paragraphs) > 7 else "示例 1:",
        "table_captions": captions[:2],
        "tables": tables[:2],
        "figure_heading": heiti_paragraphs[2] if len(heiti_paragraphs) > 2 else "图的使用示例",
        "figure_intro": songti_paragraphs[8] if len(songti_paragraphs) > 8 else "",
        "figure_caption": captions[2] if len(captions) > 2 else "",
        "figure_image_bytes": figure_image_bytes,
        "figure_image_path": str(figure_image_path) if figure_image_path else "",
        "code_heading": heiti_paragraphs[3] if len(heiti_paragraphs) > 3 else "代码的使用示例",
        "code_intro": songti_paragraphs[9] if len(songti_paragraphs) > 9 else "代码如下：",
        "code_lines": code_lines,
    }


def replace_toc_title_display(paragraph: ET.Element, new_text: str) -> None:
    runs = paragraph.findall(word_tag("r"))
    tab_index = next((idx for idx, run in enumerate(runs) if run.find(word_tag("tab")) is not None), None)
    if tab_index is None:
        return

    last_separate_index = None
    for idx, run in enumerate(runs[:tab_index]):
        fld = run.find(word_tag("fldChar"))
        fld_type = fld.get(word_attr("fldCharType")) if fld is not None else None
        if fld_type == "separate":
            last_separate_index = idx

    if last_separate_index is None:
        return

    title_text_nodes: list[ET.Element] = []
    for run in runs[last_separate_index + 1 : tab_index]:
        if run.find(word_tag("instrText")) is not None:
            continue
        title_text_nodes.extend(run.findall(word_tag("t")))

    replace_paragraph_text_nodes(title_text_nodes, new_text)


def apply_toc_sync(document_xml: str, sync_data: dict[str, object]) -> str:
    root = ET.fromstring(document_xml)
    body = root.find(word_tag("body"))
    if body is None:
        return document_xml

    toc_index = find_top_level_index(body, "目录", contains=True)
    toc_entries = [
        sync_data["full_body"]["intro_title"],
        sync_data["full_body"]["chapter_titles"][0],
        sync_data["full_body"]["section_titles"][0],
        sync_data["full_body"]["section_titles"][1],
        sync_data["full_body"]["chapter_titles"][1],
        sync_data["full_body"]["section_titles"][2],
        "4.结　论",
        "致　谢",
        sync_data["reference_title"],
        sync_data["appendix_title"],
    ]

    children = list(body)
    for offset, title in enumerate(toc_entries, start=1):
        paragraph_index = toc_index + offset
        if paragraph_index >= len(children):
            break
        if children[paragraph_index].tag != word_tag("p"):
            continue
        replace_toc_title_display(children[paragraph_index], title)

    return ET.tostring(root, encoding="unicode")


def apply_full_body_sync(document_xml: str, sync_data: dict[str, object]) -> str:
    root = ET.fromstring(document_xml)
    body = root.find(word_tag("body"))
    if body is None:
        return document_xml

    children = list(body)
    intro_idx = find_top_level_index(body, "1.引言")
    chapter2_idx = find_top_level_index(body, "2.【单击此处键入一级标题】", start=intro_idx)
    section21_idx = find_top_level_index(body, "2.1【单击此处键入二级标题】", start=chapter2_idx)
    subsection_line_idx = find_top_level_index(body, "2.1.1", start=section21_idx)
    paragraph_line_idx = find_top_level_index(body, "2.1.1.1", start=subsection_line_idx)
    section22_idx = find_top_level_index(body, "2.2【单击此处键入二级标题】", start=paragraph_line_idx)
    chapter3_idx = find_top_level_index(body, "3.【单击此处键入一级标题】", start=section22_idx)
    section31_idx = find_top_level_index(body, "3.1【单击此处键入二级标题】", start=chapter3_idx)
    formula_heading_idx = find_top_level_index(body, "公式使用示例", start=section31_idx, contains=True)
    table_heading_idx = find_top_level_index(body, "表格使用示例", start=formula_heading_idx, contains=True)
    figure_heading_idx = find_top_level_index(body, "图的使用示例", start=table_heading_idx, contains=True)
    code_heading_idx = find_top_level_index(body, "代码的使用示例", start=figure_heading_idx, contains=True)
    conclusion_idx = find_top_level_index(body, "4.结论", start=code_heading_idx, contains=True)

    children = list(body)
    reset_paragraph_text(children[intro_idx], sync_data["intro_title"], source_paragraph=children[intro_idx])
    replace_paragraph_range_with_text(
        body,
        intro_idx + 1,
        chapter2_idx,
        children[intro_idx + 1],
        sync_data["intro_paragraphs"],
        trailing_blank_count=max(0, chapter2_idx - intro_idx - 1 - len(sync_data["intro_paragraphs"])),
        blank_template=children[intro_idx + 1],
    )
    chapter2_idx = find_top_level_index(body, "2.【单击此处键入一级标题】", start=intro_idx)
    apply_format_to_top_level_paragraphs(body, intro_idx + 1, chapter2_idx, apply_body_paragraph_format)

    children = list(body)
    chapter2_idx = find_top_level_index(body, "2.【单击此处键入一级标题】", start=intro_idx)
    section21_idx = find_top_level_index(body, "2.1【单击此处键入二级标题】", start=chapter2_idx)
    subsection_line_idx = find_top_level_index(body, "2.1.1", start=section21_idx)
    paragraph_line_idx = find_top_level_index(body, "2.1.1.1", start=subsection_line_idx)
    section22_idx = find_top_level_index(body, "2.2【单击此处键入二级标题】", start=paragraph_line_idx)
    chapter3_idx = find_top_level_index(body, "3.【单击此处键入一级标题】", start=section22_idx)
    section31_idx = find_top_level_index(body, "3.1【单击此处键入二级标题】", start=chapter3_idx)
    formula_heading_idx = find_top_level_index(body, "公式使用示例", start=section31_idx, contains=True)

    reset_paragraph_text(children[chapter2_idx], sync_data["chapter_titles"][0], source_paragraph=children[chapter2_idx])
    reset_paragraph_text(children[section21_idx], sync_data["section_titles"][0], source_paragraph=children[section21_idx])
    reset_paragraph_text(children[subsection_line_idx], sync_data["subsection_line"], source_paragraph=children[subsection_line_idx])
    reset_paragraph_text(children[paragraph_line_idx], sync_data["paragraph_line"], source_paragraph=children[paragraph_line_idx])
    reset_paragraph_text(children[section22_idx], sync_data["section_titles"][1], source_paragraph=children[section22_idx])
    replace_paragraph_range_with_text(
        body,
        section22_idx + 1,
        chapter3_idx,
        children[section22_idx + 1],
        sync_data["section_body_paragraphs"][0],
        trailing_blank_count=max(0, chapter3_idx - section22_idx - 1 - len(sync_data["section_body_paragraphs"][0])),
        blank_template=children[section22_idx + 1],
    )
    chapter3_idx = find_top_level_index(body, "3.【单击此处键入一级标题】", start=section22_idx)
    apply_format_to_top_level_paragraphs(body, section22_idx + 1, chapter3_idx, apply_body_paragraph_format)

    children = list(body)
    chapter3_idx = find_top_level_index(body, "3.【单击此处键入一级标题】", start=section22_idx)
    section31_idx = find_top_level_index(body, "3.1【单击此处键入二级标题】", start=chapter3_idx)
    formula_heading_idx = find_top_level_index(body, "公式使用示例", start=section31_idx, contains=True)
    section31_body_start = next_top_level_paragraph_index(body, section31_idx + 1, formula_heading_idx)

    reset_paragraph_text(children[chapter3_idx], sync_data["chapter_titles"][1], source_paragraph=children[chapter3_idx])
    reset_paragraph_text(children[section31_idx], sync_data["section_titles"][2], source_paragraph=children[section31_idx])
    replace_paragraph_range_with_text(
        body,
        section31_body_start,
        formula_heading_idx,
        children[section31_body_start],
        sync_data["section_body_paragraphs"][1],
        trailing_blank_count=max(0, formula_heading_idx - section31_body_start - len(sync_data["section_body_paragraphs"][1])),
        blank_template=children[section31_body_start],
    )
    formula_heading_idx = find_top_level_index(body, "公式使用示例", start=section31_idx, contains=True)
    apply_format_to_top_level_paragraphs(body, section31_body_start, formula_heading_idx, apply_body_paragraph_format)

    children = list(body)
    formula_heading_idx = find_top_level_index(body, "公式使用示例", contains=True)
    table_heading_idx = find_top_level_index(body, "表格使用示例", start=formula_heading_idx, contains=True)
    figure_heading_idx = find_top_level_index(body, "图的使用示例", start=table_heading_idx, contains=True)
    code_heading_idx = find_top_level_index(body, "代码的使用示例", start=figure_heading_idx, contains=True)
    conclusion_idx = find_top_level_index(body, "4.结论", start=code_heading_idx, contains=True)

    reset_paragraph_text(children[formula_heading_idx], sync_data["formula_heading"], source_paragraph=children[formula_heading_idx])
    formula_elements = build_equation_content_elements(
        sync_data,
        children[formula_heading_idx + 1],
        children[formula_heading_idx + 2],
        children[formula_heading_idx + 3],
        children[table_heading_idx - 1],
    )
    replace_top_level_range(body, formula_heading_idx + 1, table_heading_idx, formula_elements)

    children = list(body)
    table_heading_idx = find_top_level_index(body, "表格使用示例", contains=True)
    figure_heading_idx = find_top_level_index(body, "图的使用示例", start=table_heading_idx, contains=True)
    reset_paragraph_text(children[table_heading_idx], sync_data["table_heading"], source_paragraph=children[table_heading_idx])
    reset_paragraph_text(children[table_heading_idx + 1], sync_data["table_intro"], source_paragraph=children[table_heading_idx + 1])
    reset_paragraph_text(children[table_heading_idx + 2], sync_data["table_example_label"], source_paragraph=children[table_heading_idx + 2])
    if sync_data["table_captions"]:
        reset_paragraph_text(children[table_heading_idx + 4], sync_data["table_captions"][0], source_paragraph=children[table_heading_idx + 4])
        apply_caption_paragraph_format(children[table_heading_idx + 4])
    if len(sync_data["table_captions"]) > 1:
        reset_paragraph_text(children[table_heading_idx + 7], sync_data["table_captions"][1], source_paragraph=children[table_heading_idx + 7])
        apply_caption_paragraph_format(children[table_heading_idx + 7])
    if sync_data["tables"]:
        update_table_rows(children[table_heading_idx + 5], sync_data["tables"][0])
        apply_table_format(children[table_heading_idx + 5])
    if len(sync_data["tables"]) > 1:
        update_table_rows(children[table_heading_idx + 8], sync_data["tables"][1])
        apply_table_format(children[table_heading_idx + 8])

    children = list(body)
    figure_heading_idx = find_top_level_index(body, "图的使用示例", contains=True)
    code_heading_idx = find_top_level_index(body, "代码的使用示例", start=figure_heading_idx, contains=True)
    reset_paragraph_text(children[figure_heading_idx], sync_data["figure_heading"], source_paragraph=children[figure_heading_idx])
    reset_paragraph_text(children[figure_heading_idx + 1], sync_data["figure_intro"], source_paragraph=children[figure_heading_idx + 1])
    reset_paragraph_text(children[figure_heading_idx + 3], sync_data["figure_caption"], source_paragraph=children[figure_heading_idx + 3])
    apply_caption_paragraph_format(children[figure_heading_idx + 3])
    replace_top_level_range(
        body,
        figure_heading_idx + 2,
        code_heading_idx,
        [
            copy.deepcopy(children[figure_heading_idx + 2]),
            copy.deepcopy(children[figure_heading_idx + 3]),
            make_paragraph_like(children[figure_heading_idx + 4], ""),
            make_paragraph_like(children[figure_heading_idx + 5], ""),
        ],
    )

    children = list(body)
    code_heading_idx = find_top_level_index(body, "代码的使用示例", contains=True)
    conclusion_idx = find_top_level_index(body, "4.结论", start=code_heading_idx, contains=True)
    reset_paragraph_text(children[code_heading_idx], sync_data["code_heading"], source_paragraph=children[code_heading_idx])
    reset_paragraph_text(children[code_heading_idx + 2], sync_data["code_intro"], source_paragraph=children[code_heading_idx + 2])
    code_template = children[code_heading_idx + 3]
    blank_template = children[conclusion_idx - 1]
    replace_paragraph_range_with_text(
        body,
        code_heading_idx + 3,
        conclusion_idx,
        code_template,
        sync_data["code_lines"],
        trailing_blank_count=3,
        blank_template=blank_template,
    )
    conclusion_idx = find_top_level_index(body, "4.结论", start=code_heading_idx, contains=True)
    apply_format_to_top_level_paragraphs(body, code_heading_idx + 3, conclusion_idx, apply_code_paragraph_format)
    children = list(body)
    set_page_break_before(children[conclusion_idx], True)

    return ET.tostring(root, encoding="unicode")


def build_template_sync(tex_text: str, macros: dict[str, str]) -> dict[str, object]:
    document_text = tex_text.split(r"\begin{document}", 1)[-1]

    manual_titles = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(document_text, "manualtitle")
    ]
    manual_subsections = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(document_text, "manualsubsection")
    ]
    manual_subsubsections = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(document_text, "manualsubsubsection")
    ]
    manual_paragraphs = [
        clean_tex_text(expand_tex_macros(item, macros))
        for item in extract_command_arguments(document_text, "manualparagraph")
    ]

    songti_pattern = re.compile(r"\{\\songti\\zihao\{-4\}((?:[^{}]|\{[^{}]*\})*)\\par\}", re.S)
    songti_paragraphs = [
        clean_tex_text(expand_tex_macros(match.group(1), macros))
        for match in songti_pattern.finditer(document_text)
    ]

    body_placeholder = macros.get("bodyplaceholder") or "［单击此处键入内容］"
    chapter_titles = [title for title in manual_titles if title.startswith(("2.", "3."))]
    section_titles = [title for title in manual_subsections if title.startswith(("2.1", "2.2", "3.1"))]

    body_replacements = [
        songti_paragraphs[0] if len(songti_paragraphs) > 0 else body_placeholder,
        strip_prefix(manual_subsubsections[0], "2.1.1") if manual_subsubsections else body_placeholder,
        strip_prefix(manual_paragraphs[0], "2.1.1.1") if manual_paragraphs else body_placeholder,
        songti_paragraphs[1] if len(songti_paragraphs) > 1 else body_placeholder,
        songti_paragraphs[2] if len(songti_paragraphs) > 2 else body_placeholder,
        macros.get("conclusiontext") or body_placeholder,
    ]
    body_replacements = [item or body_placeholder for item in body_replacements]

    acknowledgement_text = macros.get("acknowledgementtext") or "［单击此处键入致谢词］"
    reference_entries = [
        macros.get("referenceentryone", ""),
        macros.get("referenceentrytwo", ""),
        macros.get("referenceentrythree", ""),
        macros.get("referenceentryfour", ""),
        macros.get("referenceentryfive", ""),
        macros.get("referenceentrysix", ""),
        macros.get("referenceentryseven", ""),
        macros.get("referenceentryeight", ""),
        macros.get("referenceentrynine", ""),
        macros.get("referenceentryten", ""),
    ]

    return {
        "chapter_titles": chapter_titles,
        "section_titles": section_titles,
        "body_replacements": body_replacements,
        "acknowledgement_text": acknowledgement_text,
        "reference_title": macros.get("referencetitle") or "参考文献",
        "reference_entries": reference_entries,
        "appendix_title": macros.get("appendixtitle") or "附  录",
        "appendix_note": macros.get("appendixnote") or "（信网统一格式，供参考）",
        "appendix_a_title": macros.get("appendixatitle") or "附录A:图目录",
        "appendix_b_title": macros.get("appendixbtitle") or "附录B:表目录",
        "full_body": build_full_content_sync(tex_text, macros),
    }


def apply_document_sync(document_xml: str, sync_data: dict[str, object]) -> str:
    document_xml = apply_full_body_sync(document_xml, sync_data["full_body"])
    document_xml = apply_toc_sync(document_xml, sync_data)

    body_replacements = [safe_xml_text(item) for item in sync_data["body_replacements"]]
    document_xml = replace_sequential_occurrences(
        document_xml,
        "［单击此处键入内容］",
        body_replacements,
    )

    acknowledgement_text = safe_xml_text(sync_data["acknowledgement_text"])
    document_xml = replace_sequential_occurrences(
        document_xml,
        "［单击此处键入致谢词］",
        [acknowledgement_text],
    )

    document_xml = apply_reference_appendix_sync(document_xml, sync_data)

    return document_xml


def export_with_template(output_path: Path, macros: dict[str, str], sync_data: dict[str, object]) -> int:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    replacements = {
        placeholder: safe_xml_text(macros[macro_name])
        for placeholder, macro_name in DIRECT_TEMPLATE_REPLACEMENTS.items()
    }
    keywords = parse_keywords(
        macros["cnkeywordslineone"],
        macros["cnkeywordslinetwo"],
        macros["enkeywordslineone"],
        macros["enkeywordslinetwo"],
    )
    figure_image_bytes = sync_data["full_body"].get("figure_image_bytes")

    with zipfile.ZipFile(REFERENCE_DOCX, "r") as zin, zipfile.ZipFile(
        output_path, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith(".xml"):
                text = data.decode("utf-8")
                for old, new in replacements.items():
                    text = text.replace(old, new)
                text = replace_keyword_placeholders(text, keywords)
                if info.filename == "word/document.xml":
                    text = apply_document_sync(text, sync_data)
                data = text.encode("utf-8")
            elif info.filename == "word/media/image8.png" and figure_image_bytes:
                data = figure_image_bytes
            zout.writestr(info, data)

    print(f"已导出对齐版 Word 文件: {output_path}")
    return 0


def export_with_pandoc(output_path: Path, macros: dict[str, str]) -> int:
    pandoc = choose_pandoc()
    if pandoc is None:
        print("未找到 pandoc。", file=sys.stderr)
        print("可先执行 `bash scripts/install_pandoc.sh` 安装本地 pandoc，再重新导出。", file=sys.stderr)
        return 2

    markdown = build_markdown(macros)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmpdir:
        md_path = Path(tmpdir) / "export.md"
        md_path.write_text(markdown, encoding="utf-8")
        cmd = [
            pandoc,
            str(md_path),
            "--from",
            "markdown",
            "--to",
            "docx",
            "--reference-doc",
            str(REFERENCE_DOCX),
            "--output",
            str(output_path),
        ]
        subprocess.run(cmd, check=True)

    print(f"已导出 Pandoc 版 Word 文件: {output_path}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Word from the PPSUC LaTeX template."
    )
    parser.add_argument(
        "output",
        nargs="?",
        default=str(DEFAULT_OUTPUT),
        help="output .docx path, default: ./main.docx",
    )
    parser.add_argument(
        "--mode",
        choices=["template", "pandoc"],
        default="template",
        help="export mode: template (aligned, default) or pandoc",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = Path(args.output).resolve()

    if not MAIN_TEX.exists():
        print(f"找不到主模板文件: {MAIN_TEX}", file=sys.stderr)
        return 1
    if not REFERENCE_DOCX.exists():
        print(f"找不到官方参考 docx: {REFERENCE_DOCX}", file=sys.stderr)
        return 1

    tex_text = MAIN_TEX.read_text(encoding="utf-8")
    macros = load_macros_from_text(tex_text)
    sync_data = build_template_sync(tex_text, macros)

    try:
        if args.mode == "pandoc":
            return export_with_pandoc(output, macros)
        return export_with_template(output, macros, sync_data)
    except subprocess.CalledProcessError as exc:
        print("Pandoc 导出失败。", file=sys.stderr)
        print(f"命令: {' '.join(exc.cmd)}", file=sys.stderr)
        return exc.returncode or 1


if __name__ == "__main__":
    raise SystemExit(main())
