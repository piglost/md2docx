#!/usr/bin/env python3
"""
zhongguo-faxue-format.py — 《中国法学》格式处理器

版式以《中国法学》2026 年第 3 期《深度伪造技术全链式刑事治理模式研究》为基准：
- 页面约 541.4 x 754.0 pt
- 正标题：方正小标宋 22pt，居中
- 作者：方正楷体 14pt，居中
- 内容提要/关键词：黑体标签 + 楷体内容，PDF 实测 10.8pt，固定 18pt 行距
- 正文：方正书宋，PDF 实测 10.8pt，固定 16.5pt 行距
- 页眉：方正仿宋，PDF 实测 9.7pt，奇数页文章名，偶数页刊名
- 脚注：方正书宋，PDF 实测 7.8pt，固定 12pt 行距，段前段后 0 磅
"""

import zipfile
import tempfile
import sys
import re
import argparse
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
W = lambda tag: f"{{{W_NS}}}{tag}"
R = lambda tag: f"{{{R_NS}}}{tag}"
PR = lambda tag: f"{{{PKG_REL_NS}}}{tag}"
CT = lambda tag: f"{{{CT_NS}}}{tag}"

# 中文字号 → half-points（Word 内部单位）
SIZE = {
    "二号": "44",   # 22pt
    "小三": "30",   # 15pt
    "小二": "36",   # 18pt
    "四号": "28",   # 14pt
    "小四": "24",   # 12pt
    "五号": "21",   # 10.5pt
    "小五": "18",   # 9pt
    "实刊标题": "44",
    "实刊作者": "28",
    # Word 字号以 half-point 表示，PDF 实测 10.8pt 取最近的 11pt。
    "实刊正文": "22",
    "实刊摘要": "22",
    "实刊页眉": "19",   # PDF 实测 9.7pt，取 9.5pt。
    "实刊页码": "21",   # PDF 实测 10.3pt，取 10.5pt。
    "实刊脚注": "16",   # PDF 实测 7.8pt，取 8pt。
}

# 字体
HEI_TI = "黑体"
SONG_TI = "宋体"
SHU_SONG = "方正书宋简体"
FANG_SONG = "方正仿宋简体"
KAI_TI = "方正楷体简体"
XIAO_BIAO_SONG = "方正小标宋简体"
LANTING_CUHEI = "方正兰亭粗黑简体"
TNR = "Times New Roman"

PAGE_WIDTH = "10828"
PAGE_HEIGHT = "15080"
PAGE_MARGIN_TOP = "1440"
PAGE_MARGIN_BOTTOM = "900"
PAGE_MARGIN_LEFT = "1304"
PAGE_MARGIN_RIGHT = "1361"
BODY_FIRST_LINE = "454"
ABSTRACT_LEFT = "454"
ABSTRACT_RIGHT = "510"
ABSTRACT_FIRST_LINE = "454"
TOC_LEFT = "1032"
BODY_LINE = "330"
ABSTRACT_LINE = "360"
FOOTNOTE_LINE = "240"

# 缩进 2 字符，按 PDF 实测 10.8pt 正文字号折算，约 22.68pt。
INDENT_2CHAR = BODY_FIRST_LINE


def get_paragraph_text(p_elem: ET.Element) -> str:
    """提取段落的纯文本"""
    return "".join(t.text or "" for t in p_elem.iter(W("t"))).strip()


def get_or_create_rpr(run: ET.Element) -> ET.Element:
    """获取或创建 rPr 元素"""
    rpr = run.find(W("rPr"))
    if rpr is None:
        rpr = ET.Element(W("rPr"))
        run.insert(0, rpr)
    return rpr


def get_or_create_ppr(para: ET.Element) -> ET.Element:
    """获取或创建 pPr 元素"""
    ppr = para.find(W("pPr"))
    if ppr is None:
        ppr = ET.Element(W("pPr"))
        para.insert(0, ppr)
    return ppr


def clear_run_formatting(rpr: ET.Element, preserve_emphasis: bool = False):
    """清除期刊格式属性；正文保留加粗、斜体和下划线语义。"""
    tags = [W("rFonts"), W("sz"), W("szCs"), W("color"),
            W("highlight"), W("shd"), W("spacing"), W("kern")]
    if not preserve_emphasis:
        tags.extend([W("b"), W("bCs"), W("i"), W("iCs"), W("u")])
    for tag in tags:
        for el in rpr.findall(tag):
            rpr.remove(el)


def set_font(run: ET.Element, east_asia: str, ascii_font: str = TNR,
             size: str = SIZE["小四"], bold: bool = False,
             preserve_emphasis: bool = False):
    """设置 run 的字体、大小、加粗"""
    rpr = get_or_create_rpr(run)

    tags = [W("rFonts"), W("sz"), W("szCs"), W("color"), W("spacing")]
    if not preserve_emphasis:
        tags.extend([W("b"), W("bCs"), W("i"), W("iCs"), W("u")])
    for tag in tags:
        for old in rpr.findall(tag):
            rpr.remove(old)

    # 设置新字体
    rf = ET.SubElement(rpr, W("rFonts"))
    rf.set(W("eastAsia"), east_asia)
    rf.set(W("ascii"), ascii_font)
    rf.set(W("hAnsi"), ascii_font)
    rf.set(W("cs"), ascii_font)

    # 大小
    ET.SubElement(rpr, W("sz")).set(W("val"), size)
    ET.SubElement(rpr, W("szCs")).set(W("val"), size)

    # 加粗
    if bold:
        ET.SubElement(rpr, W("b"))
        ET.SubElement(rpr, W("bCs"))

    # 强制黑色
    color = ET.SubElement(rpr, W("color"))
    color.set(W("val"), "000000")


def set_paragraph_spacing(ppr: ET.Element, line_spacing: str = "360",
                          line_rule: str = "exact",
                          after: str = "0", before: str = "0",
                          first_line_indent: str | None = INDENT_2CHAR,
                          alignment: str | None = None,
                          preserve_indent: bool = False,
                          left: str | None = None,
                          right: str | None = None):
    """设置段落间距和对齐"""
    # 移除旧的间距设置
    tags = [W("spacing"), W("jc")]
    if not preserve_indent:
        tags.append(W("ind"))
    for tag in tags:
        for el in ppr.findall(tag):
            ppr.remove(el)

    spacing = ET.SubElement(ppr, W("spacing"))
    spacing.set(W("line"), line_spacing)
    spacing.set(W("lineRule"), line_rule)
    spacing.set(W("before"), before)
    spacing.set(W("after"), after)

    if not preserve_indent and any(
            value is not None for value in (first_line_indent, left, right)):
        ind = ET.SubElement(ppr, W("ind"))
        if first_line_indent is not None:
            ind.set(W("firstLine"), first_line_indent)
        if left is not None:
            ind.set(W("left"), left)
        if right is not None:
            ind.set(W("right"), right)

    if alignment is not None:
        jc = ET.SubElement(ppr, W("jc"))
        jc.set(W("val"), alignment)


def set_paragraph_border(ppr: ET.Element, sides: list[str],
                         size: str = "4", space: str = "4") -> None:
    for old in ppr.findall(W("pBdr")):
        ppr.remove(old)
    borders = ET.Element(W("pBdr"))
    insert_at = len(ppr)
    later_tags = {W("spacing"), W("ind"), W("jc"), W("rPr")}
    for index, child in enumerate(ppr):
        if child.tag in later_tags:
            insert_at = index
            break
    ppr.insert(insert_at, borders)
    for side in sides:
        edge = ET.SubElement(borders, W(side))
        edge.set(W("val"), "single")
        edge.set(W("sz"), size)
        edge.set(W("space"), space)
        edge.set(W("color"), "000000")


def replace_paragraph_runs(para: ET.Element, parts: list[tuple[str, str, str]]) -> None:
    """按给定文本/字体/字号重建段落 run，保留段落属性。"""
    for child in list(para):
        if child.tag != W("pPr"):
            para.remove(child)
    for text, font, size in parts:
        run = ET.SubElement(para, W("r"))
        set_font(run, font, TNR, size, bold=False)
        text_node = ET.SubElement(run, W("t"))
        if text.startswith(" ") or text.endswith(" "):
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_node.text = text


def normalize_labeled_paragraph(para: ET.Element, para_type: str) -> None:
    text = get_paragraph_text(para)
    if para_type == "abstract-label":
        match = re.match(r"^内容提要\s*[：:]?\s*(.*)$", text)
        label = "内容提要"
    else:
        match = re.match(r"^关键词\s*[：:]?\s*(.*)$", text)
        label = "关键词"
    content = match.group(1) if match else ""
    parts = [(label, HEI_TI, SIZE["实刊摘要"])]
    if content:
        parts.append((" " + content.lstrip(), KAI_TI, SIZE["实刊摘要"]))
    replace_paragraph_runs(para, parts)


def insert_toc(body: ET.Element, headings: list[tuple[int, str]]) -> None:
    """在第一个 h1 标题前插入实刊样式目次，只列一级标题。"""
    first_h1 = None
    for p in body.iter(W("p")):
        text = get_paragraph_text(p)
        if text and re.match(r"^[一二三四五六七八九十]+、", text):
            first_h1 = p
            break

    children = list(body)
    if first_h1 is not None:
        insert_pos = children.index(first_h1)
    else:
        insert_pos = 0

    h1_headings = [(level, text) for level, text in headings if level == 0]
    if not h1_headings:
        return

    toc_title = ET.Element(W("p"))
    tt_ppr = ET.SubElement(toc_title, W("pPr"))
    set_paragraph_spacing(
        tt_ppr, line_spacing=BODY_LINE, before="240", after="80",
        first_line_indent="0", alignment="center", left=TOC_LEFT)
    set_paragraph_border(tt_ppr, ["top", "left", "right"])
    tt_run = ET.SubElement(toc_title, W("r"))
    set_font(tt_run, LANTING_CUHEI, TNR, SIZE["实刊摘要"], bold=False)
    ET.SubElement(tt_run, W("t")).text = "目  次"
    body.insert(insert_pos, toc_title)
    insert_pos += 1

    for index, (_, text) in enumerate(h1_headings):
        para = ET.Element(W("p"))
        ppr = ET.SubElement(para, W("pPr"))
        set_paragraph_spacing(
            ppr, line_spacing=BODY_LINE, before="0", after="0",
            first_line_indent="0", alignment="left", left=TOC_LEFT)
        sides = ["left", "right"]
        if index == len(h1_headings) - 1:
            sides.append("bottom")
        set_paragraph_border(ppr, sides)
        run = ET.SubElement(para, W("r"))
        set_font(run, KAI_TI, TNR, SIZE["实刊摘要"], bold=False)
        ET.SubElement(run, W("t")).text = text
        body.insert(insert_pos, para)
        insert_pos += 1


def classify_paragraph(text: str, is_first: bool) -> str:
    """将段落分类为：title / author / abstract-label / abstract-content /
       keywords-label / h1 / h2 / h3 / body"""
    if not text:
        return "empty"

    if re.match(r"^内容提要\s*[：:]?", text):
        return "abstract-label"
    if re.match(r"^关键词\s*[：:]?", text):
        return "keywords-label"
    # 一级标题：一、二、三、…
    if re.match(r"^[一二三四五六七八九十]+、", text):
        return "h1"
    # 二级标题：（一）（二）…
    if re.match(r"^（[一二三四五六七八九十]+）", text):
        return "h2"
    # 三级标题：1. 2. （注意："第一，"是正文枚举标记，不是标题）
    if re.match(r"^\d+[.．、]", text):
        return "h3"

    return "body"


def classify_all_paragraphs(paragraphs: list) -> dict:
    """
    两遍扫描法分类所有段落。
    第一遍：用正则匹配锚点（title / abstract heading / h1 / h2 / h3）
    第二遍：根据锚点之间的位置关系，确定 body 段落的子类型
    """
    para_list = [(p, get_paragraph_text(p)) for p in paragraphs]
    para_list = [(p, t) for p, t in para_list if t]  # 跳过空段落

    if not para_list:
        return {}

    types = {}

    # ── 第一遍：硬匹配 ──
    for i, (p, text) in enumerate(para_list):
        pt = classify_paragraph(text, i == 0)
        if pt == "body" and i == 0:
            pt = "title"
        # 紧接着 title 的行通常是作者姓名，实刊不带“作者：”前缀。
        if i == 1 and types.get(id(para_list[0][0])) == "title" and pt == "body":
            if re.match(r"^(作者|单位|Author)", text) or re.fullmatch(r"[\u4e00-\u9fff·]{2,8}", text):
                pt = "author"
        types[id(p)] = pt

    # ── 第二遍：推断抽象区域 ──
    # 找到 abstract-label 和第一个 h1 之间的段落 → abstract-content
    abstract_start = None
    first_h1_idx = None
    for i, (p, text) in enumerate(para_list):
        pt = types.get(id(p), "body")
        if pt == "abstract-label":
            abstract_start = i
        if pt == "h1" and first_h1_idx is None:
            first_h1_idx = i

    if abstract_start is not None:
        end_idx = first_h1_idx if first_h1_idx is not None else len(para_list)
        for i in range(abstract_start + 1, end_idx):
            p, text = para_list[i]
            current = types.get(id(p), "body")
            if current in ("body", "empty"):
                types[id(p)] = "abstract-content"

    return types


def format_run(run: ET.Element, para_type: str, label_mode: bool = False):
    """按段落类型格式化单个 run"""
    preserve_emphasis = para_type in (
        "body", "abstract-content", "abstract-label", "keywords-label")
    clear_run_formatting(
        run.find(W("rPr")) if run.find(W("rPr")) is not None
        else get_or_create_rpr(run),
        preserve_emphasis=preserve_emphasis,
    )
    rpr = run.find(W("rPr"))

    if para_type == "title":
        set_font(run, XIAO_BIAO_SONG, TNR, SIZE["实刊标题"], bold=False)
    elif para_type == "author":
        set_font(run, KAI_TI, TNR, SIZE["实刊作者"], bold=False)
    elif para_type == "h1":
        set_font(run, HEI_TI, TNR, SIZE["四号"], bold=False)
    elif para_type == "h2":
        set_font(run, HEI_TI, TNR, SIZE["实刊摘要"], bold=False)
    elif para_type == "h3":
        set_font(run, SHU_SONG, TNR, SIZE["实刊正文"], bold=False)
    elif para_type in ("abstract-label", "keywords-label", "abstract-content"):
        set_font(run, KAI_TI, TNR, SIZE["实刊摘要"], bold=False,
                 preserve_emphasis=preserve_emphasis)
    else:  # body
        set_font(run, SHU_SONG, TNR, SIZE["实刊正文"], bold=False,
                 preserve_emphasis=preserve_emphasis)


def format_paragraph(para: ET.Element, para_type: str,
                     state: dict) -> None:
    """格式化整个段落"""
    ppr = get_or_create_ppr(para)

    is_list_item = ppr.find(W("numPr")) is not None

    # 列表依赖 numPr、pStyle 和原缩进，不能按普通正文清除。
    tags = [W("spacing"), W("jc"), W("keepNext"), W("keepLines")]
    if not is_list_item:
        tags.extend([W("pStyle"), W("numPr"), W("ind")])
    for tag in tags:
        for el in ppr.findall(tag):
            ppr.remove(el)

    if para_type == "title":
        set_paragraph_spacing(ppr, line_spacing="520", before="700", after="360",
                              first_line_indent="0", alignment="center")
    elif para_type == "author":
        set_paragraph_spacing(ppr, line_spacing="360", after="500",
                              first_line_indent="0", alignment="center")
    elif para_type == "h1":
        set_paragraph_spacing(ppr, line_spacing=BODY_LINE, after="220", before="360",
                              first_line_indent="0", alignment="center")
    elif para_type == "h2":
        set_paragraph_spacing(ppr, line_spacing=BODY_LINE, after="80", before="160",
                              first_line_indent="0", alignment="left")
    elif para_type == "h3":
        set_paragraph_spacing(ppr, line_spacing=BODY_LINE, after="60", before="80",
                              first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type in ("abstract-label", "keywords-label", "abstract-content"):
        set_paragraph_spacing(
            ppr, line_spacing=ABSTRACT_LINE, after="0",
            first_line_indent=ABSTRACT_FIRST_LINE, alignment="both",
            left=ABSTRACT_LEFT, right=ABSTRACT_RIGHT)
    elif para_type == "body":
        set_paragraph_spacing(ppr, line_spacing=BODY_LINE, after="0",
                              first_line_indent=(None if is_list_item else INDENT_2CHAR),
                              alignment="both", preserve_indent=is_list_item)
    else:  # empty
        pass

    # 格式化每个 run
    for run in para.findall(W("r")):
        format_run(run, para_type)

    if para_type in ("abstract-label", "keywords-label"):
        normalize_labeled_paragraph(para, para_type)


def next_relationship_ids(relationships: ET.Element, count: int) -> list[str]:
    used = []
    for rel in relationships.findall(PR("Relationship")):
        match = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if match:
            used.append(int(match.group(1)))
    start = max(used, default=0) + 1
    return [f"rId{number}" for number in range(start, start + count)]


def add_relationship(relationships: ET.Element, rel_id: str,
                     rel_type: str, target: str) -> None:
    rel = ET.SubElement(relationships, PR("Relationship"))
    rel.set("Id", rel_id)
    rel.set("Type", rel_type)
    rel.set("Target", target)


def ensure_content_type_override(content_types: ET.Element, part_name: str,
                                 content_type: str) -> None:
    for override in content_types.findall(CT("Override")):
        if override.get("PartName") == part_name:
            override.set("ContentType", content_type)
            return
    override = ET.SubElement(content_types, CT("Override"))
    override.set("PartName", part_name)
    override.set("ContentType", content_type)


def append_text_run(paragraph: ET.Element, text: str,
                    font: str = SHU_SONG,
                    size: str = SIZE["实刊页眉"]) -> ET.Element:
    run = ET.SubElement(paragraph, W("r"))
    set_font(run, font, TNR, size, bold=False)
    text_node = ET.SubElement(run, W("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def append_page_field(paragraph: ET.Element) -> None:
    begin = ET.SubElement(paragraph, W("r"))
    set_font(begin, SHU_SONG, TNR, SIZE["实刊页码"])
    ET.SubElement(begin, W("fldChar")).set(W("fldCharType"), "begin")

    instruction = ET.SubElement(paragraph, W("r"))
    set_font(instruction, SHU_SONG, TNR, SIZE["实刊页码"])
    instr_text = ET.SubElement(instruction, W("instrText"))
    instr_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_text.text = " PAGE "

    separate = ET.SubElement(paragraph, W("r"))
    set_font(separate, SHU_SONG, TNR, SIZE["实刊页码"])
    ET.SubElement(separate, W("fldChar")).set(W("fldCharType"), "separate")
    append_text_run(paragraph, "1", SHU_SONG, SIZE["实刊页码"])

    end = ET.SubElement(paragraph, W("r"))
    set_font(end, SHU_SONG, TNR, SIZE["实刊页码"])
    ET.SubElement(end, W("fldChar")).set(W("fldCharType"), "end")


def make_header(text: str, alignment: str) -> ET.ElementTree:
    root = ET.Element(W("hdr"))
    paragraph = ET.SubElement(root, W("p"))
    ppr = ET.SubElement(paragraph, W("pPr"))
    set_paragraph_spacing(
        ppr, line_spacing="240", after="0", before="0",
        first_line_indent="0", alignment=alignment)
    set_paragraph_border(ppr, ["bottom"], size="6", space="12")
    if text:
        append_text_run(paragraph, text, FANG_SONG, SIZE["实刊页眉"])
    return ET.ElementTree(root)


def make_footer(alignment: str) -> ET.ElementTree:
    root = ET.Element(W("ftr"))
    paragraph = ET.SubElement(root, W("p"))
    ppr = ET.SubElement(paragraph, W("pPr"))
    set_paragraph_spacing(
        ppr, line_spacing="240", after="0", before="0",
        first_line_indent="0", alignment=alignment)
    append_page_field(paragraph)
    return ET.ElementTree(root)


def create_header_footer_parts(tmp_path: Path, article_title: str) -> dict[str, str]:
    rels_path = tmp_path / "word" / "_rels" / "document.xml.rels"
    rels_tree = ET.parse(str(rels_path))
    relationships = rels_tree.getroot()

    targets = [
        ("header1.xml", "header"), ("header2.xml", "header"),
        ("header3.xml", "header"), ("footer1.xml", "footer"),
        ("footer2.xml", "footer"), ("footer3.xml", "footer"),
    ]
    target_names = {target for target, _ in targets}
    for rel in list(relationships.findall(PR("Relationship"))):
        if rel.get("Target") in target_names:
            relationships.remove(rel)

    rel_ids = next_relationship_ids(relationships, len(targets))
    relationship_type_base = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/")
    for rel_id, (target, kind) in zip(rel_ids, targets):
        add_relationship(relationships, rel_id, relationship_type_base + kind, target)
    ET.register_namespace("", PKG_REL_NS)
    rels_tree.write(str(rels_path), encoding="utf-8", xml_declaration=True)

    make_header(article_title, "center").write(
        str(tmp_path / "word" / "header1.xml"),
        encoding="utf-8", xml_declaration=True)
    make_header("中国法学 2026 年第3期", "center").write(
        str(tmp_path / "word" / "header2.xml"),
        encoding="utf-8", xml_declaration=True)
    make_header("", "right").write(
        str(tmp_path / "word" / "header3.xml"),
        encoding="utf-8", xml_declaration=True)
    make_footer("right").write(
        str(tmp_path / "word" / "footer1.xml"),
        encoding="utf-8", xml_declaration=True)
    make_footer("left").write(
        str(tmp_path / "word" / "footer2.xml"),
        encoding="utf-8", xml_declaration=True)
    make_footer("right").write(
        str(tmp_path / "word" / "footer3.xml"),
        encoding="utf-8", xml_declaration=True)

    content_types_path = tmp_path / "[Content_Types].xml"
    content_types_tree = ET.parse(str(content_types_path))
    content_types = content_types_tree.getroot()
    for target, kind in targets:
        ensure_content_type_override(
            content_types,
            f"/word/{target}",
            "application/vnd.openxmlformats-officedocument."
            f"wordprocessingml.{kind}+xml",
        )
    ET.register_namespace("", CT_NS)
    content_types_tree.write(
        str(content_types_path), encoding="utf-8", xml_declaration=True)

    return {
        "header_default": rel_ids[0], "header_even": rel_ids[1],
        "header_first": rel_ids[2], "footer_default": rel_ids[3],
        "footer_even": rel_ids[4], "footer_first": rel_ids[5],
    }


def configure_section(root: ET.Element, relationship_ids: dict[str, str]) -> None:
    body = root.find(W("body"))
    section = body.find(W("sectPr"))
    if section is None:
        section = ET.SubElement(body, W("sectPr"))

    for tag in (W("headerReference"), W("footerReference"), W("pgSz"),
                W("pgMar"), W("titlePg")):
        for old in section.findall(tag):
            section.remove(old)

    references = [
        ("headerReference", "default", relationship_ids["header_default"]),
        ("headerReference", "even", relationship_ids["header_even"]),
        ("headerReference", "first", relationship_ids["header_first"]),
        ("footerReference", "default", relationship_ids["footer_default"]),
        ("footerReference", "even", relationship_ids["footer_even"]),
        ("footerReference", "first", relationship_ids["footer_first"]),
    ]
    for tag, ref_type, rel_id in references:
        reference = ET.SubElement(section, W(tag))
        reference.set(W("type"), ref_type)
        reference.set(R("id"), rel_id)

    page_size = ET.SubElement(section, W("pgSz"))
    page_size.set(W("w"), PAGE_WIDTH)
    page_size.set(W("h"), PAGE_HEIGHT)
    margins = ET.SubElement(section, W("pgMar"))
    margins.set(W("top"), PAGE_MARGIN_TOP)
    margins.set(W("right"), PAGE_MARGIN_RIGHT)
    margins.set(W("bottom"), PAGE_MARGIN_BOTTOM)
    margins.set(W("left"), PAGE_MARGIN_LEFT)
    margins.set(W("header"), "700")
    margins.set(W("footer"), "500")
    margins.set(W("gutter"), "0")
    ET.SubElement(section, W("titlePg"))


def insert_before(root: ET.Element, element: ET.Element, before_tag: str) -> None:
    for index, child in enumerate(root):
        if child.tag == before_tag:
            root.insert(index, element)
            return
    root.append(element)


def configure_settings(settings_root: ET.Element) -> None:
    for old in settings_root.findall(W("evenAndOddHeaders")):
        settings_root.remove(old)
    insert_before(
        settings_root, ET.Element(W("evenAndOddHeaders")),
        W("characterSpacingControl"))
    for old in settings_root.findall(W("mirrorMargins")):
        settings_root.remove(old)
    insert_before(
        settings_root, ET.Element(W("mirrorMargins")),
        W("characterSpacingControl"))




def format_docx(input_path: str, output_path: str, add_toc: bool = False) -> int:
    """主格式化函数"""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(input_path) as zf:
            zf.extractall(tmp_path)

        doc_path = tmp_path / "word" / "document.xml"
        if not doc_path.exists():
            print("ERROR: document.xml not found")
            return 0

        tree = ET.parse(str(doc_path))
        root = tree.getroot()

        paragraphs = list(root.iter(W("p")))

        # 两遍扫描分类
        para_types = classify_all_paragraphs(paragraphs)
        article_title = next(
            (get_paragraph_text(p) for p in paragraphs
             if para_types.get(id(p)) == "title"),
            "",
        )

        # 应用格式
        state = {}
        changed = 0
        for p in paragraphs:
            pid = id(p)
            if pid in para_types:
                pt = para_types[pid]
                format_paragraph(p, pt, state)
                changed += 1

        relationship_ids = create_header_footer_parts(tmp_path, article_title)
        configure_section(root, relationship_ids)
        tree.write(str(doc_path), encoding="utf-8", xml_declaration=True)

        # 插入目次（在格式化之后，重新读取）
        if add_toc:
            # 收集标题
            level_map = {"h1": 0, "h2": 1, "h3": 2}
            headings = []
            for p in paragraphs:
                pid = id(p)
                if pid in para_types and para_types[pid] in level_map:
                    text = get_paragraph_text(p)
                    if text:
                        headings.append((level_map[para_types[pid]], text))
            # 重新解析并插入目次
            tree2 = ET.parse(str(doc_path))
            root2 = tree2.getroot()
            body2 = root2.find(W("body"))
            if body2 is not None:
                insert_toc(body2, headings)
                tree2.write(str(doc_path), encoding="utf-8", xml_declaration=True)
                print(f"   目次已插入（{len(headings)} 个标题条目）")

        settings_path = tmp_path / "word" / "settings.xml"
        if settings_path.exists():
            st_tree = ET.parse(str(settings_path))
            st_root = st_tree.getroot()
        else:
            st_root = ET.Element(W("settings"))
            st_tree = ET.ElementTree(st_root)
        configure_settings(st_root)
        st_tree.write(str(settings_path), encoding="utf-8", xml_declaration=True)

        # 同时处理脚注字体。PDF 实测脚注为方正书宋约 7.8pt，固定 12pt 行距。
        fn_path = tmp_path / "word" / "footnotes.xml"
        if fn_path.exists():
            fn_tree = ET.parse(str(fn_path))
            fn_root = fn_tree.getroot()
            for run in fn_root.iter(W("r")):
                # 编号 run 依赖 FootnoteReference 样式保持上标显示。
                if run.find(W("footnoteRef")) is not None:
                    continue
                rpr = get_or_create_rpr(run)
                clear_run_formatting(rpr, preserve_emphasis=True)
                set_font(run, SHU_SONG, TNR, SIZE["实刊脚注"], bold=False)
            # 脚注段落间距：段前段后0磅
            for p in fn_root.iter(W("p")):
                ppr = p.find(W("pPr"))
                if ppr is None:
                    ppr = ET.SubElement(p, W("pPr"))
                    p.insert(0, ppr)
                for old in ppr.findall(W("spacing")):
                    ppr.remove(old)
                sp = ET.SubElement(ppr, W("spacing"))
                sp.set(W("before"), "0")
                sp.set(W("after"), "0")
                sp.set(W("line"), FOOTNOTE_LINE)
                sp.set(W("lineRule"), "exact")
            fn_tree.write(str(fn_path), encoding="utf-8", xml_declaration=True)
            print(f"   脚注字体已统一为方正书宋 8pt")

        # 重新打包
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp_path.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp_path))

    return changed

def main():
    parser = argparse.ArgumentParser(
        description="《中国法学》格式处理器")
    parser.add_argument("input", help="输入 .docx 文件")
    parser.add_argument("output", nargs="?", default=None, help="输出 .docx 文件")
    parser.add_argument("--toc", action="store_true", help="插入自动目次")
    args = parser.parse_args()

    input_path = args.input
    output_path = args.output or str(
        Path(input_path).parent / (Path(input_path).stem + "_中国法学格式.docx"))

    print(f"📐 《中国法学》格式化...")
    changed = format_docx(input_path, output_path, add_toc=args.toc)
    print(f"   格式化段落: {changed}")
    print(f"✅ 输出: {output_path}")


if __name__ == "__main__":
    main()
