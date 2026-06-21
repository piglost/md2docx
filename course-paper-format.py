#!/usr/bin/env python3
"""
course-paper-format.py — 课程论文格式处理器

版式以《论公共数据开放的数据价值实现原则.docx》OOXML 提取结果为基准：
- 页面 A4，纵向；上下 72pt，左右 90pt
- 标题：Heiti SC Light 18pt，加粗，居中
- 摘要/关键词：标签 Songti SC 10.5pt 加粗，正文仿宋 10.5pt，左右缩进 24pt
- 正文：Songti SC 12pt，两端对齐，首行缩进 24pt，1.5 倍行距
- 一级标题：Heiti SC Medium 14pt，加粗，居中
- 二级标题：Heiti SC Medium 12pt，加粗，左对齐
- 脚注：宋体 9pt，段前段后 0，紧凑行距
"""

import argparse
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)
W = lambda tag: f"{{{W_NS}}}{tag}"

TNR = "Times New Roman"
HEITI_LIGHT = "Heiti SC Light"
HEITI_MEDIUM = "Heiti SC Medium"
SONGTI_SC = "Songti SC"
FANGSONG = "仿宋"
SONGTI = "宋体"
HELVETICA = "Helvetica"

PAGE_WIDTH = "11906"
PAGE_HEIGHT = "16838"
MARGIN_TOP_BOTTOM = "1440"
MARGIN_LEFT_RIGHT = "1800"
HEADER_DISTANCE = "851"
FOOTER_DISTANCE = "992"

SIZE_TITLE = "36"
SIZE_BODY = "24"
SIZE_H1 = "28"
SIZE_H2 = "24"
SIZE_ABSTRACT = "21"
SIZE_FOOTNOTE = "18"

BODY_LINE = "360"
BODY_FIRST_LINE = "480"
ABSTRACT_SIDE_INDENT = "480"
ABSTRACT_FIRST_LINE = "420"
FOOTNOTE_LINE = "240"


def get_text(paragraph: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph.iter(W("t"))).strip()


def get_or_create_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find(W("pPr"))
    if ppr is None:
        ppr = ET.Element(W("pPr"))
        paragraph.insert(0, ppr)
    return ppr


def get_or_create_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find(W("rPr"))
    if rpr is None:
        rpr = ET.Element(W("rPr"))
        run.insert(0, rpr)
    return rpr


def clear_children(parent: ET.Element, tags: tuple[str, ...]) -> None:
    for tag in tags:
        for child in parent.findall(tag):
            parent.remove(child)


def set_font(run: ET.Element, east_asia: str, ascii_font: str = TNR,
             size: str = SIZE_BODY, bold: bool = False,
             preserve_emphasis: bool = False) -> None:
    rpr = get_or_create_rpr(run)
    tags = [W("rFonts"), W("sz"), W("szCs"), W("color"), W("highlight"), W("shd")]
    if not preserve_emphasis:
        tags.extend([W("b"), W("bCs"), W("i"), W("iCs"), W("u")])
    clear_children(rpr, tuple(tags))

    fonts = ET.SubElement(rpr, W("rFonts"))
    fonts.set(W("eastAsia"), east_asia)
    fonts.set(W("ascii"), ascii_font)
    fonts.set(W("hAnsi"), ascii_font)
    fonts.set(W("cs"), ascii_font)

    ET.SubElement(rpr, W("sz")).set(W("val"), size)
    ET.SubElement(rpr, W("szCs")).set(W("val"), size)
    if bold:
        ET.SubElement(rpr, W("b"))
        ET.SubElement(rpr, W("bCs"))

    ET.SubElement(rpr, W("color")).set(W("val"), "000000")


def set_paragraph(paragraph: ET.Element, line: str = BODY_LINE,
                  line_rule: str = "auto", before: str = "0",
                  after: str = "0", first_line: str | None = BODY_FIRST_LINE,
                  alignment: str = "both", left: str | None = None,
                  right: str | None = None, keep: bool = False) -> None:
    ppr = get_or_create_ppr(paragraph)
    clear_children(
        ppr,
        (W("pStyle"), W("numPr"), W("spacing"), W("ind"), W("jc"),
         W("keepNext"), W("keepLines"), W("pageBreakBefore")),
    )

    if keep:
        ET.SubElement(ppr, W("keepNext"))
        ET.SubElement(ppr, W("keepLines"))

    spacing = ET.SubElement(ppr, W("spacing"))
    spacing.set(W("line"), line)
    spacing.set(W("lineRule"), line_rule)
    spacing.set(W("before"), before)
    spacing.set(W("after"), after)

    if any(value is not None for value in (first_line, left, right)):
        indent = ET.SubElement(ppr, W("ind"))
        if first_line is not None:
            indent.set(W("firstLine"), first_line)
        if left is not None:
            indent.set(W("left"), left)
        if right is not None:
            indent.set(W("right"), right)

    ET.SubElement(ppr, W("jc")).set(W("val"), alignment)


def replace_runs(paragraph: ET.Element,
                 parts: list[tuple[str, str, str, bool]]) -> None:
    for child in list(paragraph):
        if child.tag != W("pPr"):
            paragraph.remove(child)
    for text, font, size, bold in parts:
        run = ET.SubElement(paragraph, W("r"))
        set_font(run, font, TNR, size, bold=bold)
        node = ET.SubElement(run, W("t"))
        if text.startswith(" ") or text.endswith(" "):
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        node.text = text


def normalize_labeled_paragraph(paragraph: ET.Element, label: str) -> None:
    text = get_text(paragraph)
    pattern = rf"^{label}\s*[：:]?\s*(.*)$"
    match = re.match(pattern, text)
    content = match.group(1) if match else ""
    parts = [(f"{label}：", SONGTI_SC, SIZE_ABSTRACT, True)]
    if content:
        parts.append((" " + content.lstrip(), FANGSONG, SIZE_ABSTRACT, False))
    replace_runs(paragraph, parts)


def is_intro_heading(text: str) -> bool:
    return re.sub(r"\s+", "", text) == "引言"


def classify_paragraph(text: str, index: int) -> str:
    if not text:
        return "empty"
    if index == 0:
        return "title"
    if index == 1 and re.fullmatch(r"[\u4e00-\u9fffA-Za-z·\s]{2,20}", text):
        return "author"
    if re.match(r"^摘要\s*[：:]?", text):
        return "abstract"
    if re.match(r"^关键词\s*[：:]?", text):
        return "keywords"
    if re.match(r"^[一二三四五六七八九十]+、", text):
        return "h1"
    if re.match(r"^（[一二三四五六七八九十]+）", text):
        return "h2"
    if re.match(r"^\d+[.．、]\s*", text):
        return "h2"
    return "body"


def classify_all(paragraphs: list[ET.Element]) -> dict[int, str]:
    non_empty = [(p, get_text(p)) for p in paragraphs]
    non_empty = [(p, text) for p, text in non_empty if text]
    types: dict[int, str] = {}
    seen_section_heading = False
    for i, (paragraph, text) in enumerate(non_empty):
        para_type = classify_paragraph(text, i)
        if para_type == "body" and is_intro_heading(text) and not seen_section_heading:
            para_type = "h1"
        if para_type in ("h1", "h2"):
            seen_section_heading = True
        types[id(paragraph)] = para_type
    return types


def format_paragraph(paragraph: ET.Element, para_type: str) -> None:
    if para_type == "title":
        set_paragraph(paragraph, before="156", after="156",
                      first_line="0", alignment="center", keep=True)
        font, size, bold = HEITI_LIGHT, SIZE_TITLE, True
    elif para_type == "author":
        set_paragraph(paragraph, after="156", first_line="0", alignment="center")
        font, size, bold = SONGTI_SC, SIZE_BODY, False
    elif para_type == "h1":
        set_paragraph(paragraph, before="156", after="156",
                      first_line="0", alignment="center", keep=True)
        font, size, bold = HEITI_MEDIUM, SIZE_H1, True
    elif para_type == "h2":
        set_paragraph(paragraph, before="80", after="60",
                      first_line="0", alignment="left", keep=True)
        font, size, bold = HEITI_MEDIUM, SIZE_H2, True
    elif para_type in ("abstract", "keywords"):
        set_paragraph(
            paragraph, first_line=ABSTRACT_FIRST_LINE, alignment="both",
            left=ABSTRACT_SIDE_INDENT, right=ABSTRACT_SIDE_INDENT, keep=True)
        normalize_labeled_paragraph(
            paragraph, "摘要" if para_type == "abstract" else "关键词")
        return
    else:
        set_paragraph(paragraph)
        font, size, bold = SONGTI_SC, SIZE_BODY, False

    preserve_emphasis = para_type == "body"
    for run in paragraph.findall(W("r")):
        set_font(run, font, TNR, size, bold=bold,
                 preserve_emphasis=preserve_emphasis)


def configure_section(root: ET.Element) -> None:
    body = root.find(W("body"))
    if body is None:
        return
    section = body.find(W("sectPr"))
    if section is None:
        section = ET.SubElement(body, W("sectPr"))

    clear_children(section, (W("pgSz"), W("pgMar"), W("cols")))
    page_size = ET.SubElement(section, W("pgSz"))
    page_size.set(W("w"), PAGE_WIDTH)
    page_size.set(W("h"), PAGE_HEIGHT)

    margins = ET.SubElement(section, W("pgMar"))
    margins.set(W("top"), MARGIN_TOP_BOTTOM)
    margins.set(W("right"), MARGIN_LEFT_RIGHT)
    margins.set(W("bottom"), MARGIN_TOP_BOTTOM)
    margins.set(W("left"), MARGIN_LEFT_RIGHT)
    margins.set(W("header"), HEADER_DISTANCE)
    margins.set(W("footer"), FOOTER_DISTANCE)
    margins.set(W("gutter"), "0")

    ET.SubElement(section, W("cols")).set(W("space"), "425")


def format_footnotes(tmp_path: Path) -> None:
    footnotes_path = tmp_path / "word" / "footnotes.xml"
    if not footnotes_path.exists():
        return
    tree = ET.parse(str(footnotes_path))
    root = tree.getroot()
    for paragraph in root.iter(W("p")):
        set_paragraph(paragraph, line=FOOTNOTE_LINE, before="0", after="0",
                      first_line="360", alignment="left")
    for run in root.iter(W("r")):
        if run.find(W("footnoteRef")) is not None:
            continue
        set_font(run, SONGTI, HELVETICA, SIZE_FOOTNOTE,
                 preserve_emphasis=True)
    tree.write(str(footnotes_path), encoding="utf-8", xml_declaration=True)


def format_docx(input_path: str, output_path: str) -> int:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(input_path) as archive:
            archive.extractall(tmp_path)

        document_path = tmp_path / "word" / "document.xml"
        if not document_path.exists():
            print("ERROR: document.xml not found", file=sys.stderr)
            return 0

        tree = ET.parse(str(document_path))
        root = tree.getroot()
        paragraphs = list(root.iter(W("p")))
        para_types = classify_all(paragraphs)

        changed = 0
        for paragraph in paragraphs:
            para_type = para_types.get(id(paragraph))
            if para_type:
                format_paragraph(paragraph, para_type)
                changed += 1

        configure_section(root)
        tree.write(str(document_path), encoding="utf-8", xml_declaration=True)
        format_footnotes(tmp_path)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for path in tmp_path.rglob("*"):
                if path.is_file():
                    archive.write(path, path.relative_to(tmp_path))

    return changed


def main() -> None:
    parser = argparse.ArgumentParser(description="课程论文格式处理器")
    parser.add_argument("input", help="输入 .docx 文件")
    parser.add_argument("output", nargs="?", default=None, help="输出 .docx 文件")
    args = parser.parse_args()

    output_path = args.output or str(
        Path(args.input).parent / (Path(args.input).stem + "_课程论文格式.docx"))
    print("📐 课程论文格式化...")
    changed = format_docx(args.input, output_path)
    print(f"   格式化段落: {changed}")
    print(f"✅ 输出: {output_path}")


if __name__ == "__main__":
    main()
