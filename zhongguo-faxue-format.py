#!/usr/bin/env python3
"""
zhongguo-faxue-format.py — 《中国法学》格式处理器

按照《中国法学》投稿格式要求，对 docx 进行精确格式化：
- 正标题：宋体，二号(22pt)，加粗，居中
- 作者信息：宋体，小三(15pt)，居中
- 内容提要/关键词：宋体，小四(12pt)，与正文一致
- 一级标题：黑体，四号(14pt)，居中，不加粗
- 二级标题：黑体，小四(12pt)，两端对齐，不加粗
- 三级标题：宋体，小四(12pt)，不加粗
- 正文：宋体，小四(12pt)，首行缩进2字符
- 正文行距：1.5倍行距(360 twips)
- 脚注：宋体，小五(9pt)
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

# 中文字号 → half-points（Word 内部单位）
SIZE = {
    "二号": "44",   # 22pt
    "小三": "30",   # 15pt
    "小二": "36",   # 18pt
    "四号": "28",   # 14pt
    "小四": "24",   # 12pt
    "五号": "21",   # 10.5pt
    "小五": "18",   # 9pt
}

# 字体
HEI_TI = "黑体"
SONG_TI = "宋体"
FANG_SONG = "仿宋"
TNR = "Times New Roman"

# 缩进 2 字符 ≈ 480 twips（按五号字体宽度估算，精确值可调整）
INDENT_2CHAR = "480"


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


def clear_run_formatting(rpr: ET.Element):
    """清除 run 级别的格式"""
    for tag in [W("rFonts"), W("sz"), W("szCs"), W("b"), W("bCs"),
                W("color"), W("highlight"), W("shd"), W("i"), W("iCs"),
                W("u"), W("spacing"), W("kern")]:
        for el in rpr.findall(tag):
            rpr.remove(el)


def set_font(run: ET.Element, east_asia: str, ascii_font: str = TNR,
             size: str = SIZE["小四"], bold: bool = False):
    """设置 run 的字体、大小、加粗"""
    rpr = get_or_create_rpr(run)

    # 移除旧的字体设置
    old_rf = rpr.find(W("rFonts"))
    if old_rf is not None:
        rpr.remove(old_rf)

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
                          after: str = "0", before: str = "0",
                          first_line_indent: str | None = INDENT_2CHAR,
                          alignment: str | None = None):
    """设置段落间距和对齐"""
    # 移除旧的间距设置
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
        if first_line_indent == "0":
            ind.set(W("firstLine"), "0")
        else:
            ind.set(W("firstLine"), first_line_indent)

    if alignment is not None:
        jc = ET.SubElement(ppr, W("jc"))
        jc.set(W("val"), alignment)


def insert_toc(body: ET.Element, headings: list[tuple[int, str]]) -> None:
    """在第一个 h1 标题前插入纯文本目次，带矩形边框。

    参数：
        body: Word 文档 body 元素
        headings: [(level, text), ...] level 0=h1, 1=h2, 2=h3
    """
    # 找到第一个 h1 段落在 body 中的位置
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

    # ── "目 次" 标题段落（带边框）──
    toc_title = ET.Element(W("p"))
    tt_ppr = ET.SubElement(toc_title, W("pPr"))
    tt_jc = ET.SubElement(tt_ppr, W("jc"))
    tt_jc.set(W("val"), "center")
    tt_spacing = ET.SubElement(tt_ppr, W("spacing"))
    tt_spacing.set(W("before"), "240")
    tt_spacing.set(W("after"), "120")
    # 边框：仅上、左、右（标题和内容共享边框视觉效果）
    tt_pbdr = ET.SubElement(tt_ppr, W("pBdr"))
    for side in ["top", "left", "right"]:
        el = ET.SubElement(tt_pbdr, W(side))
        el.set(W("val"), "single")
        el.set(W("sz"), "4")
        el.set(W("space"), "4")
        el.set(W("color"), "000000")
    tt_run = ET.SubElement(toc_title, W("r"))
    tt_rpr = ET.SubElement(tt_run, W("rPr"))
    tt_rf = ET.SubElement(tt_rpr, W("rFonts"))
    tt_rf.set(W("eastAsia"), HEI_TI)
    tt_rf.set(W("ascii"), TNR)
    tt_rf.set(W("hAnsi"), TNR)
    ET.SubElement(tt_rpr, W("sz")).set(W("val"), SIZE["四号"])
    ET.SubElement(tt_rpr, W("szCs")).set(W("val"), SIZE["四号"])
    ET.SubElement(tt_run, W("t")).text = "目  次"
    body.insert(insert_pos, toc_title)
    insert_pos += 1

    # ── 各条目段落 ──
    # 缩进：h1=0, h2=2字符, h3=4字符
    indent_map = {0: "0", 1: INDENT_2CHAR, 2: str(int(INDENT_2CHAR) * 2)}
    for level, text in headings:
        para = ET.Element(W("p"))
        ppr = ET.SubElement(para, W("pPr"))
        # 左缩进
        ind = ET.SubElement(ppr, W("ind"))
        indent_val = indent_map.get(level, "0")
        ind.set(W("left"), indent_val)
        # 间距
        sp = ET.SubElement(ppr, W("spacing"))
        sp.set(W("line"), "300")
        sp.set(W("lineRule"), "auto")
        sp.set(W("before"), "0")
        sp.set(W("after"), "0")
        # 左侧边框 + 上下边框（与标题段落共同构成矩形）
        pbdr = ET.SubElement(ppr, W("pBdr"))
        for side in ["left", "bottom", "right"]:
            el = ET.SubElement(pbdr, W(side))
            el.set(W("val"), "single")
            el.set(W("sz"), "4")
            el.set(W("space"), "4")
            el.set(W("color"), "000000")
        # 文本 run
        run = ET.SubElement(para, W("r"))
        rpr = ET.SubElement(run, W("rPr"))
        rf = ET.SubElement(rpr, W("rFonts"))
        rf.set(W("eastAsia"), SONG_TI)
        rf.set(W("ascii"), TNR)
        rf.set(W("hAnsi"), TNR)
        ET.SubElement(rpr, W("sz")).set(W("val"), SIZE["五号"])
        ET.SubElement(rpr, W("szCs")).set(W("val"), SIZE["五号"])
        color = ET.SubElement(rpr, W("color"))
        color.set(W("val"), "000000")
        t = ET.SubElement(run, W("t"))
        t.text = text
        body.insert(insert_pos, para)
        insert_pos += 1


def classify_paragraph(text: str, is_first: bool) -> str:
    """将段落分类为：title / author / abstract-label / abstract-content /
       keywords-label / h1 / h2 / h3 / body"""
    if not text:
        return "empty"

    # 注意：内容提要/摘要/关键词 不做特殊分类，统一按 body 处理
    # （PDF原文中这些与正文字体字号一致，均为宋体小四）
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
        # 紧接着 title 的行是作者信息
        if i == 1 and types.get(id(para_list[0][0])) == "title" and pt == "body":
            if re.match(r"^(作者|单位|Author)", text):
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
    clear_run_formatting(run.find(W("rPr")) if run.find(W("rPr")) is not None
                         else get_or_create_rpr(run))
    rpr = run.find(W("rPr"))

    if para_type == "title":
        set_font(run, SONG_TI, TNR, SIZE["二号"], bold=True)
    elif para_type == "author":
        set_font(run, SONG_TI, TNR, SIZE["小三"], bold=False)
    elif para_type == "h1":
        set_font(run, HEI_TI, TNR, SIZE["四号"], bold=False)
    elif para_type == "h2":
        set_font(run, HEI_TI, TNR, SIZE["小四"], bold=False)
    elif para_type == "h3":
        set_font(run, SONG_TI, TNR, SIZE["小四"], bold=False)
    else:  # body
        set_font(run, SONG_TI, TNR, SIZE["小四"], bold=False)


def format_paragraph(para: ET.Element, para_type: str,
                     state: dict) -> None:
    """格式化整个段落"""
    ppr = get_or_create_ppr(para)

    # 清除旧的段落级样式
    for tag in [W("pStyle"), W("numPr"), W("spacing"), W("ind"), W("jc"),
                W("keepNext"), W("keepLines")]:
        for el in ppr.findall(tag):
            ppr.remove(el)

    if para_type == "title":
        set_paragraph_spacing(ppr, line_spacing="360", after="200",
                              first_line_indent="0", alignment="center")
    elif para_type == "author":
        set_paragraph_spacing(ppr, line_spacing="360", after="120",
                              first_line_indent="0", alignment="center")
    elif para_type == "h1":
        set_paragraph_spacing(ppr, line_spacing="360", after="120", before="240",
                              first_line_indent="0", alignment="center")
    elif para_type == "h2":
        set_paragraph_spacing(ppr, line_spacing="360", after="80", before="160",
                              first_line_indent="0", alignment="both")
    elif para_type == "h3":
        set_paragraph_spacing(ppr, line_spacing="360", after="60", before="80",
                              first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "body":
        set_paragraph_spacing(ppr, line_spacing="360", after="0",
                              first_line_indent=INDENT_2CHAR, alignment="both")
    else:  # empty
        pass

    # 格式化每个 run
    for run in para.findall(W("r")):
        format_run(run, para_type)




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

        # 应用格式
        state = {}
        changed = 0
        for p in paragraphs:
            pid = id(p)
            if pid in para_types:
                pt = para_types[pid]
                format_paragraph(p, pt, state)
                changed += 1

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

        # 同时处理脚注字体（小五 宋体）
        fn_path = tmp_path / "word" / "footnotes.xml"
        if fn_path.exists():
            fn_tree = ET.parse(str(fn_path))
            fn_root = fn_tree.getroot()
            for run in fn_root.iter(W("r")):
                set_font(run, SONG_TI, TNR, SIZE["小五"], bold=False)
                rpr = run.find(W("rPr"))
                if rpr is not None:
                    # 移除脚注引用样式
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
            print(f"   脚注字体已统一为宋体小五")

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
