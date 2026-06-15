#!/usr/bin/env python3
"""
zhongguo-sheke-format.py — 《中国社会科学》格式处理器

按照《中国社会科学》投稿格式要求，对 docx 进行精确格式化：
- 全文字体统一为宋体（仅大标题加粗，其余均不加粗）
- 正标题：宋体，二号(22pt)，加粗，居中
- 副标题：宋体，小三(15pt)，居中，不加粗
- 摘要/关键词：宋体，小四(12pt)，首行缩进2字符，不加粗
- 一级标题：宋体，四号(14pt)，居中，不加粗（"引言"中间空3字符）
- 二级标题：宋体，小四(12pt)，两端对齐，首行缩进2字符
- 三级标题：宋体，小四(12pt)，两端对齐
- 正文：宋体，小四(12pt)，首行缩进2字符
- 正文行距：1.5倍行距(360 twips)
- 脚注：仿宋，五号(10.5pt)，每页重新编号
"""

import zipfile
import tempfile
import sys
import re
import argparse
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)
W = lambda tag: f"{{{W_NS}}}{tag}"

SIZE = {
    "二号": "44", "三号": "32", "小三": "30", "小二": "36",
    "四号": "28", "小四": "24", "五号": "21", "小五": "18",
}

HEI_TI = "黑体"
SONG_TI = "宋体"
FANG_SONG = "仿宋"
TNR = "Times New Roman"
INDENT_2CHAR = "480"


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


def clear_run_formatting(rpr):
    for tag in [W("rFonts"), W("sz"), W("szCs"), W("b"), W("bCs"),
                W("color"), W("highlight"), W("shd"), W("i"), W("iCs"),
                W("u"), W("spacing"), W("kern")]:
        for el in rpr.findall(tag):
            rpr.remove(el)


def set_font(run, east_asia, ascii_font=TNR, size=SIZE["小四"], bold=False):
    rpr = get_or_create_rpr(run)
    old_rf = rpr.find(W("rFonts"))
    if old_rf is not None:
        rpr.remove(old_rf)
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
    color = ET.SubElement(rpr, W("color"))
    color.set(W("val"), "000000")


def set_paragraph_spacing(ppr, line_spacing="360", after="0", before="0",
                          first_line_indent=INDENT_2CHAR, alignment=None):
    for tag in [W("spacing"), W("ind"), W("jc")]:
        for el in ppr.findall(tag):
            ppr.remove(el)
    spacing = ET.SubElement(ppr, W("spacing"))
    spacing.set(W("line"), line_spacing)
    spacing.set(W("lineRule"), "auto")
    spacing.set(W("before"), before)
    spacing.set(W("after"), after)
    if first_line_indent is not None:
        ind = ET.SubElement(ppr, W("ind"))
        ind.set(W("firstLine"), first_line_indent if first_line_indent != "0" else "0")
    if alignment is not None:
        jc = ET.SubElement(ppr, W("jc"))
        jc.set(W("val"), alignment)


def classify_paragraph(text, is_first):
    if not text:
        return "empty"
    if text.strip() in ("摘要", "内容提要", "内容摘要"):
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

    # 作者单位：keywords 后、abstract-label 前的含括号 body 段落
    keywords_idx = None
    abstract_idx = None
    for i, (p, text) in enumerate(para_list):
        pt = types.get(id(p), "body")
        if pt == "keywords-label":
            keywords_idx = i
        if pt == "abstract-label":
            abstract_idx = i
            break
    if keywords_idx is not None and abstract_idx is not None:
        for i in range(keywords_idx + 1, abstract_idx):
            p, text = para_list[i]
            if types.get(id(p), "body") == "body" and re.search(r"（.*?）", text):
                types[id(p)] = "author-unit"

    return types


def format_run(run, para_type):
    clear_run_formatting(run.find(W("rPr")) if run.find(W("rPr")) is not None
                         else get_or_create_rpr(run))
    # 全部用宋体，仅 title 加粗
    bold = para_type == "title"
    if para_type == "author-unit":
        set_font(run, SONG_TI, TNR, SIZE["小五"], bold=False)
    elif para_type in ("h1", "h2", "h3"):
        size = SIZE["四号"] if para_type == "h1" else SIZE["小四"]
        set_font(run, SONG_TI, TNR, size, bold=False)
    elif para_type == "subtitle":
        set_font(run, SONG_TI, TNR, SIZE["小三"], bold=False)
    elif para_type == "author":
        set_font(run, SONG_TI, TNR, SIZE["三号"], bold=False)
    elif para_type == "title":
        set_font(run, SONG_TI, TNR, SIZE["二号"], bold=True)
    else:
        set_font(run, SONG_TI, TNR, SIZE["小四"], bold=bold)


def format_paragraph(para, para_type):
    ppr = get_or_create_ppr(para)
    for tag in [W("pStyle"), W("numPr"), W("spacing"), W("ind"), W("jc"),
                W("keepNext"), W("keepLines")]:
        for el in ppr.findall(tag):
            ppr.remove(el)

    if para_type == "title":
        set_paragraph_spacing(ppr, after="200", first_line_indent="0", alignment="center")
    elif para_type in ("subtitle", "author"):
        set_paragraph_spacing(ppr, after="120", first_line_indent="0", alignment="center")
    elif para_type == "abstract-label":
        set_paragraph_spacing(ppr, after="0", first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "keywords-label":
        set_paragraph_spacing(ppr, after="0", first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "author-unit":
        set_paragraph_spacing(ppr, after="0", first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "h1":
        set_paragraph_spacing(ppr, after="120", before="240",
                              first_line_indent="0", alignment="center")
    elif para_type == "h2":
        set_paragraph_spacing(ppr, after="80", before="160",
                              first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "h3":
        set_paragraph_spacing(ppr, after="60", before="80",
                              first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "body":
        set_paragraph_spacing(ppr, after="0", first_line_indent=INDENT_2CHAR, alignment="both")

    for run in para.findall(W("r")):
        format_run(run, para_type)

    # "引言" → "引   言"
    if para_type == "h1":
        for run in para.findall(W("r")):
            for t in run.findall(W("t")):
                if t.text and t.text.strip() == "引言":
                    t.text = "引   言"


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

        changed = 0
        for p in paragraphs:
            pid = id(p)
            if pid in para_types:
                format_paragraph(p, para_types[pid])
                changed += 1

        tree.write(str(doc_path), encoding="utf-8", xml_declaration=True)

        # 脚注：仿宋五号 + 每页重新编号
        fn_path = tmp_path / "word" / "footnotes.xml"
        if fn_path.exists():
            fn_tree = ET.parse(str(fn_path))
            fn_root = fn_tree.getroot()
            for run in fn_root.iter(W("r")):
                set_font(run, FANG_SONG, TNR, SIZE["五号"], bold=False)
                rpr = run.find(W("rPr"))
                if rpr is not None:
                    rstyle = rpr.find(W("rStyle"))
                    if rstyle is not None:
                        rpr.remove(rstyle)
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
                sp.set(W("line"), "240")
                sp.set(W("lineRule"), "auto")
            fn_tree.write(str(fn_path), encoding="utf-8", xml_declaration=True)
            print("   脚注字体已统一为仿宋五号")

        # 设置脚注每页重新编号
        settings_path = tmp_path / "word" / "settings.xml"
        if settings_path.exists():
            st_tree = ET.parse(str(settings_path))
            st_root = st_tree.getroot()
        else:
            st_root = ET.Element(W("settings"))
            st_tree = ET.ElementTree(st_root)
        for old in st_root.findall(W("footnotePr")):
            st_root.remove(old)
        fn_pr = ET.SubElement(st_root, W("footnotePr"))
        fn_pr.set(W("pos"), "pageBottom")
        restart = ET.SubElement(fn_pr, W("numRestart"))
        restart.set(W("val"), "eachPage")
        fmt = ET.SubElement(fn_pr, W("numFmt"))
        fmt.set(W("val"), "decimal")
        st_tree.write(str(settings_path), encoding="utf-8", xml_declaration=True)
        print("   脚注每页重新编号已设置")

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
