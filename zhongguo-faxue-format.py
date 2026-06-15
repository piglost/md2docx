#!/usr/bin/env python3
"""
zhongguo-faxue-format.py — 《中国法学》格式处理器

按照《中国法学》投稿格式要求，对 docx 进行精确格式化：
- 正标题：黑体，小二(18pt)，加粗，两端对齐
- 摘要/关键词标签：宋体，五号(10.5pt)，加粗；内容：仿宋，五号，不加粗，左右缩进2字符
- 一级标题：黑体，四号(14pt)，居中，不加粗
- 二级标题：黑体，小四(12pt)，两端对齐，不加粗
- 三级标题：宋体，小四(12pt)，不加粗
- 正文：宋体，小四(12pt)，首行缩进2字符
- 脚注：宋体，小五(9pt)
"""

import zipfile
import tempfile
import sys
import re
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)
W = lambda tag: f"{{{W_NS}}}{tag}"

# 中文字号 → half-points（Word 内部单位）
SIZE = {
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


def classify_paragraph(text: str, is_first: bool) -> str:
    """将段落分类为：title / author / abstract-label / abstract-content /
       keywords-label / h1 / h2 / h3 / body"""
    if not text:
        return "empty"

    # 摘要标签（含独立"摘要"）
    if re.match(r"^(内容提要|内容摘要|摘要)[：:]?\s*$", text):
        return "abstract-label"
    if re.match(r"^(内容提要|内容摘要|摘要)[：:]", text):
        return "abstract-label"
    if text.strip() in ("摘要", "内容提要", "内容摘要"):
        return "abstract-label"

    # 关键词标签
    if re.match(r"^关键词[：:]", text):
        return "keywords-label"

    # 一级标题：一、二、三、…
    if re.match(r"^[一二三四五六七八九十]+、", text):
        return "h1"
    # 二级标题：（一）（二）…
    if re.match(r"^（[一二三四五六七八九十]+）", text):
        return "h2"
    # 三级标题：1. 2. 或 第一，
    if re.match(r"^\d+[.．、]", text):
        return "h3"
    if re.match(r"^第[一二三四五六七八九十]+[，,]", text):
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
        set_font(run, HEI_TI, TNR, SIZE["小二"], bold=True)
    elif para_type == "abstract-label":
        set_font(run, SONG_TI, TNR, SIZE["五号"], bold=True)
    elif para_type == "abstract-content":
        set_font(run, FANG_SONG, TNR, SIZE["五号"], bold=False)
    elif para_type == "keywords-label":
        set_font(run, SONG_TI, TNR, SIZE["五号"], bold=True)
    elif para_type == "author":
        set_font(run, SONG_TI, TNR, SIZE["小四"], bold=False)
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
    elif para_type == "abstract-label":
        set_paragraph_spacing(ppr, line_spacing="300", after="0",
                              first_line_indent="0", alignment="both")
    elif para_type == "abstract-content":
        # 左右缩进 2 字符
        ind = ppr.find(W("ind"))
        if ind is None:
            ind = ET.SubElement(ppr, W("ind"))
        ind.set(W("left"), INDENT_2CHAR)
        ind.set(W("right"), INDENT_2CHAR)
        ind.set(W("firstLine"), INDENT_2CHAR)
        set_paragraph_spacing(ppr, line_spacing="300", after="0",
                              first_line_indent=INDENT_2CHAR, alignment="both")
    elif para_type == "keywords-label":
        set_paragraph_spacing(ppr, line_spacing="300", after="0",
                              first_line_indent="0", alignment="both")
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

    # 特殊处理：摘要和关键词标签后的内容
    if para_type == "abstract-label":
        state["in_abstract"] = True
        state["in_keywords"] = False
    elif para_type == "keywords-label":
        state["in_abstract"] = False
        state["in_keywords"] = True


def format_docx(input_path: str, output_path: str) -> int:
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
            fn_tree.write(str(fn_path), encoding="utf-8", xml_declaration=True)
            print(f"   脚注字体已统一为宋体小五")

        # 重新打包
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp_path.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp_path))

    return changed


def main():
    if len(sys.argv) < 2:
        print("用法: python3 zhongguo-faxue-format.py <input.docx> [output.docx]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else \
        str(Path(input_path).parent / (Path(input_path).stem + "_中国法学格式.docx"))

    print(f"📐 《中国法学》格式化...")
    changed = format_docx(input_path, output_path)
    print(f"   格式化段落: {changed}")
    print(f"✅ 输出: {output_path}")


if __name__ == "__main__":
    main()
