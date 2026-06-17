#!/usr/bin/env python3
"""
zhongguo-sheke-format.py — 《中国社会科学》格式处理器

版式以《中国社会科学》2026 年第 2 期实刊为基准：
- 页面约 201 x 280 mm，正文版心约 438 pt
- 正标题：方正小标宋 26pt，不加粗
- 作者：方正楷体 15pt，疏排
- 摘要/关键词：黑体标签 + 仿宋内容，10pt
- 正文：方正书宋 11pt，固定 18.6pt 行距
- 脚注：方正仿宋 10pt，圆圈数字，每页重新编号
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

SIZE = {
    "二号": "44", "三号": "32", "小三": "30", "小二": "36",
    "四号": "28", "小四": "24", "五号": "21", "小五": "18",
    "实刊标题": "52", "实刊作者": "30", "实刊正文": "22",
    "实刊摘要": "20", "实刊脚注": "20",
}

HEI_TI = "方正黑体简体"
SONG_TI = "方正书宋简体"
FANG_SONG = "方正仿宋简体"
KAI_TI = "方正楷体简体"
XIAO_BIAO_SONG = "方正小标宋简体"
TNR = "Times New Roman"
INDENT_2CHAR = "480"

PAGE_WIDTH = "11396"
PAGE_HEIGHT = "15874"
PAGE_MARGIN_TOP = "1440"
PAGE_MARGIN_BOTTOM = "1800"
PAGE_MARGIN_LEFT = "1538"
PAGE_MARGIN_RIGHT = "1090"
ABSTRACT_SIDE_INDENT = "600"
ABSTRACT_FIRST_LINE = "420"
BODY_LINE = "372"
ABSTRACT_LINE = "338"
FOOTNOTE_LINE = "280"


def get_paragraph_text(p_elem):
    return "".join(t.text or "" for t in p_elem.iter(W("t"))).strip()


def get_or_create_rpr(run):
    rpr = run.find(W("rPr"))
    if rpr is None:
        rpr = ET.Element(W("rPr"))
        run.insert(0, rpr)
    return rpr


def get_or_create_ppr(para):
    ppr = para.find(W("pPr"))
    if ppr is None:
        ppr = ET.Element(W("pPr"))
        para.insert(0, ppr)
    return ppr


def clear_run_formatting(rpr, preserve_emphasis=False):
    tags = [W("rFonts"), W("sz"), W("szCs"), W("color"),
            W("highlight"), W("shd"), W("spacing"), W("kern")]
    if not preserve_emphasis:
        tags.extend([W("b"), W("bCs"), W("i"), W("iCs"), W("u")])
    for tag in tags:
        for el in rpr.findall(tag):
            rpr.remove(el)


def set_font(run, east_asia, ascii_font=TNR, size=SIZE["实刊正文"],
             bold=False, character_spacing=None):
    rpr = get_or_create_rpr(run)
    for tag in [W("rFonts"), W("sz"), W("szCs"), W("b"), W("bCs"),
                W("color"), W("spacing")]:
        for old in rpr.findall(tag):
            rpr.remove(old)
    rf = ET.SubElement(rpr, W("rFonts"))
    rf.set(W("eastAsia"), east_asia)
    rf.set(W("ascii"), ascii_font)
    rf.set(W("hAnsi"), ascii_font)
    rf.set(W("cs"), ascii_font)
    ET.SubElement(rpr, W("sz")).set(W("val"), size)
    ET.SubElement(rpr, W("szCs")).set(W("val"), size)
    if bold:
        ET.SubElement(rpr, W("b"))
        ET.SubElement(rpr, W("bCs"))
    if character_spacing is not None:
        ET.SubElement(rpr, W("spacing")).set(W("val"), character_spacing)
    color = ET.SubElement(rpr, W("color"))
    color.set(W("val"), "000000")


def set_paragraph_spacing(ppr, line_spacing=BODY_LINE, line_rule="exact",
                          after="0", before="0",
                          first_line_indent=INDENT_2CHAR, alignment=None,
                          preserve_indent=False, left=None, right=None):
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


def set_bottom_border(ppr):
    for old in ppr.findall(W("pBdr")):
        ppr.remove(old)
    borders = ET.Element(W("pBdr"))
    # pBdr 在 OOXML 段落属性中必须位于 spacing/ind/jc 之前。
    insert_at = len(ppr)
    later_tags = {W("spacing"), W("ind"), W("jc"), W("rPr")}
    for index, child in enumerate(ppr):
        if child.tag in later_tags:
            insert_at = index
            break
    ppr.insert(insert_at, borders)
    bottom = ET.SubElement(borders, W("bottom"))
    bottom.set(W("val"), "single")
    bottom.set(W("sz"), "6")
    bottom.set(W("space"), "12")
    bottom.set(W("color"), "000000")


def replace_paragraph_runs(para, parts):
    """用指定的文本与格式重建段落 run，保留 pPr 和书签。"""
    for child in list(para):
        if child.tag not in (W("pPr"), W("bookmarkStart"), W("bookmarkEnd")):
            para.remove(child)
    for text, font, size, bold in parts:
        run = ET.SubElement(para, W("r"))
        set_font(run, font, TNR, size, bold=bold)
        text_node = ET.SubElement(run, W("t"))
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_node.text = text


def normalize_labeled_paragraph(para, para_type):
    text = get_paragraph_text(para)
    if para_type == "abstract-label":
        match = re.match(r"^摘\s*要\s*[：:]?\s*(.*)$", text)
        content = match.group(1) if match else ""
        label = "摘  要" + ("：" if content else "")
    else:
        match = re.match(r"^关键词\s*[：:]?\s*(.*)$", text)
        content = match.group(1) if match else ""
        label = "关键词" + ("：" if content else "")
    parts = [(label, HEI_TI, SIZE["实刊摘要"], False)]
    if content:
        parts.append((content, FANG_SONG, SIZE["实刊摘要"], False))
    replace_paragraph_runs(para, parts)


def classify_paragraph(text, is_first):
    if not text:
        return "empty"
    if re.fullmatch(r"摘\s*要", text.strip()) or text.strip() in ("内容提要", "内容摘要"):
        return "abstract-label"
    if re.match(r"^摘\s*要[：:]", text):
        return "abstract-label"
    if re.match(r"^关键词[：:]", text):
        return "keywords-label"
    if re.match(r"^[一二三四五六七八九十]+、", text):
        return "h1"
    if text.strip() in ("引言", "引  言", "结语", "结  语", "结论", "结  论"):
        return "h1"
    if re.match(r"^（[一二三四五六七八九十]+）", text):
        return "h2"
    if re.match(r"^\d+[.．、]", text):
        return "h3"
    return "body"


def classify_all_paragraphs(paragraphs):
    para_list = [(p, get_paragraph_text(p)) for p in paragraphs]
    para_list = [(p, t) for p, t in para_list if t]
    if not para_list:
        return {}

    types = {}
    for i, (p, text) in enumerate(para_list):
        pt = classify_paragraph(text, i == 0)
        if pt == "body" and i == 0:
            pt = "title"
        if i == 1 and types.get(id(para_list[0][0])) == "title" and pt == "body":
            if text.startswith("——") or text.startswith("—"):
                pt = "subtitle"
        if i <= 2 and pt == "body":
            prev_types = [types.get(id(para_list[j][0])) for j in range(i)]
            if "title" in prev_types and re.match(r"^[\u4e00-\u9fff]{2,4}$", text.replace(" ", "")):
                pt = "author"
        types[id(p)] = pt

    # 独立摘要标题后的正文，直到关键词之前，属于摘要内容。
    abstract_idx = None
    keywords_idx = None
    for i, (p, text) in enumerate(para_list):
        pt = types.get(id(p), "body")
        if pt == "abstract-label" and abstract_idx is None:
            abstract_idx = i
        if pt == "keywords-label":
            keywords_idx = i
            break

    if abstract_idx is not None and keywords_idx is not None:
        label_text = para_list[abstract_idx][1]
        if not re.match(r"^摘\s*要\s*[：:].+", label_text):
            for i in range(abstract_idx + 1, keywords_idx):
                p, _ = para_list[i]
                if types.get(id(p)) == "body":
                    types[id(p)] = "abstract-content"

    # 关键词之后、首个普通正文之前的作者说明行。
    if keywords_idx is not None:
        for i in range(keywords_idx + 1, min(keywords_idx + 4, len(para_list))):
            p, text = para_list[i]
            if types.get(id(p), "body") != "body":
                continue
            if text.startswith("作者") or (
                    re.search(r"（.*?\d+.*?）", text)
                    and any(word in text for word in ("大学", "研究所", "学院", "教授"))):
                types[id(p)] = "author-unit"
                break

    return types


def format_run(run, para_type):
    preserve_emphasis = para_type in (
        "body", "abstract-content", "author-unit")
    clear_run_formatting(
        run.find(W("rPr")) if run.find(W("rPr")) is not None
        else get_or_create_rpr(run),
        preserve_emphasis=preserve_emphasis,
    )
    if para_type == "author-unit":
        set_font(run, FANG_SONG, TNR, SIZE["实刊摘要"], bold=False)
    elif para_type == "abstract-content":
        set_font(run, FANG_SONG, TNR, SIZE["实刊摘要"], bold=False)
    elif para_type == "h1":
        set_font(run, XIAO_BIAO_SONG, TNR, SIZE["实刊作者"], bold=False)
    elif para_type in ("h2", "h3"):
        set_font(run, SONG_TI, TNR, SIZE["实刊正文"], bold=False)
    elif para_type == "subtitle":
        set_font(run, KAI_TI, TNR, SIZE["小二"], bold=False)
    elif para_type == "author":
        set_font(run, KAI_TI, TNR, SIZE["实刊作者"], bold=False,
                 character_spacing="180")
    elif para_type == "title":
        set_font(run, XIAO_BIAO_SONG, TNR, SIZE["实刊标题"], bold=False)
    else:
        set_font(run, SONG_TI, TNR, SIZE["实刊正文"], bold=False)


def format_paragraph(para, para_type):
    ppr = get_or_create_ppr(para)
    is_list_item = ppr.find(W("numPr")) is not None
    tags = [W("spacing"), W("jc"), W("keepNext"), W("keepLines")]
    if not is_list_item:
        tags.extend([W("pStyle"), W("numPr"), W("ind")])
    for tag in tags:
        for el in ppr.findall(tag):
            ppr.remove(el)

    if para_type == "title":
        set_paragraph_spacing(
            ppr, line_spacing="620", before="500", after="360",
            first_line_indent="0", alignment="center")
    elif para_type == "subtitle":
        set_paragraph_spacing(
            ppr, line_spacing="440", after="240",
            first_line_indent="0", alignment="center")
    elif para_type == "author":
        set_paragraph_spacing(
            ppr, line_spacing="360", after="440",
            first_line_indent="0", alignment="center")
    elif para_type == "abstract-label":
        set_paragraph_spacing(
            ppr, line_spacing=ABSTRACT_LINE, after="0",
            first_line_indent=ABSTRACT_FIRST_LINE, alignment="both",
            left=ABSTRACT_SIDE_INDENT, right=ABSTRACT_SIDE_INDENT)
    elif para_type == "abstract-content":
        set_paragraph_spacing(
            ppr, line_spacing=ABSTRACT_LINE, after="0",
            first_line_indent=ABSTRACT_FIRST_LINE, alignment="both",
            left=ABSTRACT_SIDE_INDENT, right=ABSTRACT_SIDE_INDENT)
    elif para_type == "keywords-label":
        set_paragraph_spacing(
            ppr, line_spacing=ABSTRACT_LINE, before="80", after="0",
            first_line_indent=ABSTRACT_FIRST_LINE, alignment="both",
            left=ABSTRACT_SIDE_INDENT, right=ABSTRACT_SIDE_INDENT)
    elif para_type == "author-unit":
        set_paragraph_spacing(
            ppr, line_spacing=ABSTRACT_LINE, before="240", after="500",
            first_line_indent=ABSTRACT_FIRST_LINE, alignment="both",
            left=ABSTRACT_SIDE_INDENT, right=ABSTRACT_SIDE_INDENT)
        set_bottom_border(ppr)
    elif para_type == "h1":
        set_paragraph_spacing(ppr, line_spacing="420", after="240", before="360",
                              first_line_indent="0", alignment="center")
    elif para_type == "h2":
        set_paragraph_spacing(ppr, after="120", before="240",
                              first_line_indent="0", alignment="left",
                              left=ABSTRACT_SIDE_INDENT)
    elif para_type == "h3":
        set_paragraph_spacing(ppr, after="80", before="160",
                              first_line_indent="0", alignment="left",
                              left="420")
    elif para_type == "body":
        set_paragraph_spacing(
            ppr, after="0",
            first_line_indent=(None if is_list_item else INDENT_2CHAR),
            alignment="both", preserve_indent=is_list_item)

    for run in para.findall(W("r")):
        format_run(run, para_type)

    if para_type in ("abstract-label", "keywords-label"):
        normalize_labeled_paragraph(para, para_type)

    # "引言" → "引   言"
    if para_type == "h1":
        for run in para.findall(W("r")):
            for t in run.findall(W("t")):
                if t.text and t.text.strip() == "引言":
                    t.text = "引   言"


def next_relationship_ids(relationships, count):
    used = []
    for rel in relationships.findall(PR("Relationship")):
        match = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if match:
            used.append(int(match.group(1)))
    start = max(used, default=0) + 1
    return [f"rId{number}" for number in range(start, start + count)]


def add_relationship(relationships, rel_id, rel_type, target):
    rel = ET.SubElement(relationships, PR("Relationship"))
    rel.set("Id", rel_id)
    rel.set("Type", rel_type)
    rel.set("Target", target)


def ensure_content_type_override(content_types, part_name, content_type):
    for override in content_types.findall(CT("Override")):
        if override.get("PartName") == part_name:
            override.set("ContentType", content_type)
            return
    override = ET.SubElement(content_types, CT("Override"))
    override.set("PartName", part_name)
    override.set("ContentType", content_type)


def append_text_run(paragraph, text, font=SONG_TI,
                    size=SIZE["实刊摘要"]):
    run = ET.SubElement(paragraph, W("r"))
    set_font(run, font, TNR, size, bold=False)
    text_node = ET.SubElement(run, W("t"))
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def append_page_field(paragraph):
    begin = ET.SubElement(paragraph, W("r"))
    set_font(begin, SONG_TI, TNR, SIZE["实刊摘要"])
    ET.SubElement(begin, W("fldChar")).set(W("fldCharType"), "begin")

    instruction = ET.SubElement(paragraph, W("r"))
    set_font(instruction, SONG_TI, TNR, SIZE["实刊摘要"])
    instr_text = ET.SubElement(instruction, W("instrText"))
    instr_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_text.text = " PAGE "

    separate = ET.SubElement(paragraph, W("r"))
    set_font(separate, SONG_TI, TNR, SIZE["实刊摘要"])
    ET.SubElement(separate, W("fldChar")).set(W("fldCharType"), "separate")
    append_text_run(paragraph, "1")

    end = ET.SubElement(paragraph, W("r"))
    set_font(end, SONG_TI, TNR, SIZE["实刊摘要"])
    ET.SubElement(end, W("fldChar")).set(W("fldCharType"), "end")


def make_header(text, alignment="center", show_rule=True):
    root = ET.Element(W("hdr"))
    paragraph = ET.SubElement(root, W("p"))
    ppr = ET.SubElement(paragraph, W("pPr"))
    set_paragraph_spacing(
        ppr, line_spacing="240", first_line_indent="0", alignment=alignment)
    if show_rule:
        set_bottom_border(ppr)
    if text:
        append_text_run(paragraph, text)
    return ET.ElementTree(root)


def make_footer(alignment):
    root = ET.Element(W("ftr"))
    paragraph = ET.SubElement(root, W("p"))
    ppr = ET.SubElement(paragraph, W("pPr"))
    set_paragraph_spacing(
        ppr, line_spacing="240", first_line_indent="0", alignment=alignment)
    append_text_run(paragraph, "· ")
    append_page_field(paragraph)
    append_text_run(paragraph, " ·")
    return ET.ElementTree(root)


def create_header_footer_parts(tmp_path, article_title):
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
        add_relationship(
            relationships, rel_id, relationship_type_base + kind, target)
    # Office 包关系文件使用默认命名空间；LibreOffice 不接受 ns0 前缀形式。
    ET.register_namespace("", PKG_REL_NS)
    rels_tree.write(str(rels_path), encoding="utf-8", xml_declaration=True)

    make_header(article_title, alignment="right").write(
        str(tmp_path / "word" / "header1.xml"),
        encoding="utf-8", xml_declaration=True)
    make_header("中国社会科学", alignment="left").write(
        str(tmp_path / "word" / "header2.xml"),
        encoding="utf-8", xml_declaration=True)
    make_header("", show_rule=False).write(
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


def configure_section(root, relationship_ids):
    body = root.find(W("body"))
    section = body.find(W("sectPr"))
    if section is None:
        section = ET.SubElement(body, W("sectPr"))

    for tag in (W("headerReference"), W("footerReference"), W("pgSz"),
                W("pgMar"), W("titlePg"), W("footnotePr")):
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

    footnote_properties = ET.SubElement(section, W("footnotePr"))
    ET.SubElement(footnote_properties, W("pos")).set(W("val"), "pageBottom")
    ET.SubElement(footnote_properties, W("numFmt")).set(
        W("val"), "decimalEnclosedCircle")
    ET.SubElement(footnote_properties, W("numRestart")).set(
        W("val"), "eachPage")

    page_size = ET.SubElement(section, W("pgSz"))
    page_size.set(W("w"), PAGE_WIDTH)
    page_size.set(W("h"), PAGE_HEIGHT)
    margins = ET.SubElement(section, W("pgMar"))
    margins.set(W("top"), PAGE_MARGIN_TOP)
    margins.set(W("right"), PAGE_MARGIN_RIGHT)
    margins.set(W("bottom"), PAGE_MARGIN_BOTTOM)
    margins.set(W("left"), PAGE_MARGIN_LEFT)
    margins.set(W("header"), "700")
    margins.set(W("footer"), "600")
    margins.set(W("gutter"), "0")
    ET.SubElement(section, W("titlePg"))


def insert_before(root, element, before_tag):
    for index, child in enumerate(root):
        if child.tag == before_tag:
            root.insert(index, element)
            return
    root.append(element)


def configure_settings(settings_root):
    for tag in (W("evenAndOddHeaders"), W("mirrorMargins"), W("footnotePr")):
        for old in settings_root.findall(tag):
            settings_root.remove(old)
    insert_before(
        settings_root, ET.Element(W("mirrorMargins")), W("proofState"))
    insert_before(
        settings_root, ET.Element(W("evenAndOddHeaders")),
        W("characterSpacingControl"))
    footnote_properties = ET.Element(W("footnotePr"))
    ET.SubElement(footnote_properties, W("pos")).set(W("val"), "pageBottom")
    ET.SubElement(footnote_properties, W("numFmt")).set(
        W("val"), "decimalEnclosedCircle")
    ET.SubElement(footnote_properties, W("numRestart")).set(
        W("val"), "eachPage")
    insert_before(settings_root, footnote_properties, W("rsids"))


def format_docx(input_path, output_path):
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
        para_types = classify_all_paragraphs(paragraphs)
        article_title = next(
            (get_paragraph_text(p) for p in paragraphs
             if para_types.get(id(p)) == "title"),
            "",
        )

        changed = 0
        for p in paragraphs:
            pid = id(p)
            if pid in para_types:
                format_paragraph(p, para_types[pid])
                changed += 1

        relationship_ids = create_header_footer_parts(tmp_path, article_title)
        configure_section(root, relationship_ids)
        tree.write(str(doc_path), encoding="utf-8", xml_declaration=True)

        # 脚注：仿宋 10pt，固定行距，段前段后均为 0 磅。
        fn_path = tmp_path / "word" / "footnotes.xml"
        if fn_path.exists():
            fn_tree = ET.parse(str(fn_path))
            fn_root = fn_tree.getroot()
            for run in fn_root.iter(W("r")):
                if run.find(W("footnoteRef")) is not None:
                    continue
                rpr = get_or_create_rpr(run)
                clear_run_formatting(rpr, preserve_emphasis=True)
                set_font(run, FANG_SONG, TNR, SIZE["实刊脚注"], bold=False)
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
            print("   脚注字体已统一为方正仿宋 10pt")

        # 设置镜像版心、奇偶页页眉和圆圈脚注编号。
        settings_path = tmp_path / "word" / "settings.xml"
        if settings_path.exists():
            st_tree = ET.parse(str(settings_path))
            st_root = st_tree.getroot()
        else:
            st_root = ET.Element(W("settings"))
            st_tree = ET.ElementTree(st_root)
        configure_settings(st_root)
        st_tree.write(str(settings_path), encoding="utf-8", xml_declaration=True)
        print("   实刊页面、页眉页码和圆圈脚注编号已设置")

        # 重新打包
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp_path.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp_path))

    return changed


def main():
    parser = argparse.ArgumentParser(description="《中国社会科学》格式处理器")
    parser.add_argument("input", help="输入 .docx 文件")
    parser.add_argument("output", nargs="?", default=None, help="输出 .docx 文件")
    args = parser.parse_args()

    input_path = args.input
    output_path = args.output or str(
        Path(input_path).parent / (Path(input_path).stem + "_中国社会科学格式.docx"))

    print(f"📐 《中国社会科学》格式化...")
    changed = format_docx(input_path, output_path)
    print(f"   格式化段落: {changed}")
    print(f"✅ 输出: {output_path}")


if __name__ == "__main__":
    main()
