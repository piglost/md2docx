#!/usr/bin/env python3
"""
md2docx.py — Markdown to Word 转换器（法律写作专用）

功能：
1. 将 [1][2] 等文内引用标记转换为 Word 脚注（footnotes）
2. 将文末「参考文献」条目转为脚注定义
3. 输出纯黑文字（无颜色）

用法：
    python3 md2docx.py input.md [output.docx] [--reference-doc template.docx]

依赖：
    pandoc >= 3.0（brew install pandoc）
"""

import re
import sys
import os
import subprocess
import tempfile
import zipfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
STRIP_COLOR_FILTER = SCRIPT_DIR / "strip-color.lua"

MAX_CITATION_RANGE = 100


class Md2DocxError(RuntimeError):
    """md2docx 可预期错误的基类。"""


class CitationError(Md2DocxError):
    """引用或参考文献格式错误。"""


class ConversionError(Md2DocxError):
    """外部转换或 DOCX 后处理失败。"""

# legal-homework-formatter 的颜色修复脚本（第二道防线）
FIX_HEADING_SCRIPT = (
    Path.home() / ".claude/skills/legal-homework-formatter/scripts/"
    "fix_pandoc_heading_artifacts.py"
)

def normalize_font_to_songti(docx_path: str) -> bool:
    """
    将 docx 文件中所有文字的字体统一设为宋体（中文）+ Times New Roman（英文/数字）。
    直接修改 XML，不依赖 python-docx。
    """
    import zipfile
    from xml.etree import ElementTree as ET

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ET.register_namespace("w", W_NS)
    W = lambda tag: f"{{{W_NS}}}{tag}"

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(docx_path, 'r') as zf:
            zf.extractall(tmp)

        document_xml = tmp / "word" / "document.xml"
        if not document_xml.exists():
            return False

        tree = ET.parse(str(document_xml))
        root = tree.getroot()

        changed = 0
        for run in root.iter(W("r")):
            rpr = run.find(W("rPr"))
            if rpr is None:
                rpr = ET.SubElement(run, W("rPr"))
                run.insert(0, rpr)

            # 移除已有的 rFonts
            existing_fonts = rpr.find(W("rFonts"))
            if existing_fonts is not None:
                rpr.remove(existing_fonts)

            # 设置字体：中文宋体，英文 Times New Roman
            rf = ET.SubElement(rpr, W("rFonts"))
            rf.set(W("ascii"), "Times New Roman")
            rf.set(W("hAnsi"), "Times New Roman")
            rf.set(W("eastAsia"), "宋体")
            rf.set(W("cs"), "Times New Roman")
            changed += 1

        tree.write(str(document_xml), encoding="utf-8", xml_declaration=True)

        # 重新打包
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp))

        return changed > 0

# ── 配置 ──────────────────────────────────────────────
REF_SECTION_TITLES = [
    "参考文献", "参考文献与注释", "参考书目",
    "References", "Bibliography",
    "注释", "尾注",
]

# 文内引用模式。匹配范围不包含引用后的空白。
INLINE_CITE_RE = re.compile(r"\[(\d+(?:\s*[,，]\s*\d+|\s*[-–—]\s*\d+)*)\]")

# 参考文献条目：支持两种格式
# 格式1：[1] 作者. 标题...（行首带方括号）
# 格式2：1. 作者. 标题...（行首数字+英文句号）
REF_ENTRY_BRACKET_RE = re.compile(r"^\s*\[(\d+)\]\s*(.*)")
REF_ENTRY_DOT_RE = re.compile(r"^\s*(\d+)\.\s+(.*)")
# 参考文献内部的子标题
REF_SUBHEADING_RE = re.compile(r"^#{1,4}\s+")
# 圆圈数字脚注标记 ①②③...⑳
CIRCLED_NUM_RE = re.compile(r"[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")
CIRCLED_NUM_MAP = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
# 全角方括号数字引用 〔1〕〔2〕
FULLWIDTH_BRACKET_RE = re.compile(r"〔(\d+)〕")
REFERENCE_ENTRY_PATTERNS = (
    ("bracket", REF_ENTRY_BRACKET_RE),
    ("dot", REF_ENTRY_DOT_RE),
    ("fullwidth", re.compile(r"^\s*〔(\d+)〕\s*(.*)")),
    ("circled", re.compile(r"^\s*([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])\s*(.*)")),
)
FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")


def _circled_to_number(value: str) -> str:
    return str(CIRCLED_NUM_MAP.index(value) + 1)


def _match_reference_entry(line: str):
    for style, pattern in REFERENCE_ENTRY_PATTERNS:
        match = pattern.match(line)
        if not match:
            continue
        if style == "circled":
            return _circled_to_number(match.group(1)), match.group(2), style
        return match.group(1), match.group(2), style
    return None


def _store_reference(ref_dict: dict[str, str], key: str | None,
                     parts: list[str]) -> None:
    if key is None:
        return
    if key in ref_dict:
        raise CitationError(f"参考文献编号重复: {key}")
    content = " ".join(part for part in parts if part).strip()
    if not content:
        raise CitationError(f"参考文献 {key} 内容为空")
    ref_dict[key] = content


def parse_reference_section(lines: list[str]) -> tuple[list[str], dict[str, str], set[str]]:
    """解析参考文献区，统一支持四种编号格式和多行续行。"""
    boundary = find_ref_section_boundary(lines)
    if boundary is None:
        return lines, {}, set()

    body_lines = lines[:boundary]
    ref_lines = lines[boundary + 1:]
    ref_dict: dict[str, str] = {}
    styles: set[str] = set()
    current_key: str | None = None
    current_parts: list[str] = []
    unparsed: list[str] = []

    for line in ref_lines:
        stripped = line.strip()
        if REF_SUBHEADING_RE.match(stripped):
            _store_reference(ref_dict, current_key, current_parts)
            current_key, current_parts = None, []
            continue

        entry = _match_reference_entry(stripped)
        if entry:
            _store_reference(ref_dict, current_key, current_parts)
            current_key, content, style = entry
            current_parts = [content.strip()] if content.strip() else []
            styles.add(style)
            continue

        if not stripped:
            continue
        if current_key is not None:
            current_parts.append(stripped)
        else:
            unparsed.append(stripped)

    _store_reference(ref_dict, current_key, current_parts)
    if not ref_dict and unparsed:
        raise CitationError("参考文献章节存在内容，但没有识别到受支持的编号条目")
    return body_lines, ref_dict, styles


def find_circled_ref_boundary(lines: list[str]) -> int | None:
    """找到注释 section（①②③ 风格脚注定义所在位置）"""
    for title in ["注释", "注  释"]:
        heading_pat = re.compile(rf"^#+\s+{re.escape(title)}\s*$")
        for i, line in enumerate(lines):
            if heading_pat.match(line.strip()):
                return i
    return None


def split_circled_footnotes(lines: list[str]) -> tuple[list[str], dict[str, str]]:
    """将文档分成正文和圆圈数字注释部分。
    返回：(正文行, {数字str: 注释内容})
    """
    body, refs, styles = parse_reference_section(lines)
    if "circled" not in styles:
        return lines, {}
    return body, refs


def convert_circled_citations(text: str, ref_dict: dict[str, str]) -> str:
    """将 ①②③ 标记转为 pandoc 脚注 [^N]"""
    if not ref_dict:
        return text
    return _replace_citations_in_text(text, ref_dict, {"circled"})


def find_ref_section_boundary(lines: list[str]) -> int | None:
    """
    找到参考文献 section 的起始行索引。
    返回 None 表示没有专门的参考文献节。
    """
    title_pattern = "|".join(re.escape(title) for title in REF_SECTION_TITLES)
    heading_pattern = re.compile(rf"^#{{1,4}}\s+(?:{title_pattern})\s*$")
    for index, line in enumerate(lines):
        if heading_pattern.match(line.strip()):
            return index
    return None


def split_document(lines: list[str]) -> tuple[list[str], dict[str, str]]:
    """
    将文档分成正文部分和参考文献部分。
    返回：(正文行列表, {编号: 参考文献内容})
    """
    body, refs, _ = parse_reference_section(lines)
    return body, refs


def _expand_citation(value: str) -> list[str]:
    normalized = re.sub(r"\s+", "", value)
    if re.search(r"[-–—]", normalized):
        parts = re.split(r"[-–—]", normalized)
        if len(parts) != 2:
            raise CitationError(f"无效引用范围: [{value}]")
        start, end = map(int, parts)
        if start > end:
            raise CitationError(f"引用范围起点大于终点: [{value}]")
        if end - start + 1 > MAX_CITATION_RANGE:
            raise CitationError(
                f"引用范围过大: [{value}]，最多允许 {MAX_CITATION_RANGE} 项")
        return [str(number) for number in range(start, end + 1)]
    return [part for part in re.split(r"[,，]", normalized) if part]


def _validate_reference_ids(ids: list[str], ref_dict: dict[str, str],
                            original: str) -> None:
    missing = [number for number in ids if number not in ref_dict]
    if missing:
        raise CitationError(
            f"引用 {original} 缺少参考文献定义: {', '.join(missing)}")


def _consume_markdown_link(text: str, index: int) -> int | None:
    """若当前位置是 Markdown 链接或图片，返回其结束位置。"""
    label_start = index
    is_image = False
    if text.startswith("![", index):
        label_start = index + 1
        is_image = True
    elif text[index:index + 1] != "[":
        return None

    label_end = text.find("]", label_start + 1)
    if label_end == -1:
        return None
    suffix_start = label_end + 1

    if text[suffix_start:suffix_start + 1] == "(":
        depth = 1
        cursor = suffix_start + 1
        while cursor < len(text) and depth:
            if text[cursor] == "\\":
                cursor += 2
                continue
            if text[cursor] == "(":
                depth += 1
            elif text[cursor] == ")":
                depth -= 1
            cursor += 1
        return cursor if depth == 0 else None

    if text[suffix_start:suffix_start + 1] == "[":
        reference_end = text.find("]", suffix_start + 1)
        if reference_end == -1:
            return None
        first_label = text[label_start:label_end + 1]
        second_label = text[suffix_start:reference_end + 1]
        # [1][2] 是连续脚注，不是引用式链接。
        if (not is_image and INLINE_CITE_RE.fullmatch(first_label)
                and INLINE_CITE_RE.fullmatch(second_label)):
            return None
        return reference_end + 1
    return None


def _replace_citations_in_text(text: str, ref_dict: dict[str, str],
                               styles: set[str]) -> str:
    output: list[str] = []
    index = 0
    while index < len(text):
        if text[index] == "\\" and index + 1 < len(text):
            output.append(text[index:index + 2])
            index += 2
            continue

        if text[index] == "`":
            end_ticks = index
            while end_ticks < len(text) and text[end_ticks] == "`":
                end_ticks += 1
            delimiter = text[index:end_ticks]
            closing = text.find(delimiter, end_ticks)
            if closing == -1:
                output.append(text[index:])
                break
            closing += len(delimiter)
            output.append(text[index:closing])
            index = closing
            continue

        link_end = _consume_markdown_link(text, index)
        if link_end is not None:
            output.append(text[index:link_end])
            index = link_end
            continue

        if text[index] == "<":
            tag_end = text.find(">", index + 1)
            if tag_end != -1:
                output.append(text[index:tag_end + 1])
                index = tag_end + 1
                continue

        match = (INLINE_CITE_RE.match(text, index)
                 if styles.intersection({"bracket", "dot"}) else None)
        if match:
            ids = _expand_citation(match.group(1))
            _validate_reference_ids(ids, ref_dict, match.group(0))
            output.append("".join(f"[^{number}]" for number in dict.fromkeys(ids)))
            index = match.end()
            continue

        if "fullwidth" in styles:
            fullwidth_match = FULLWIDTH_BRACKET_RE.match(text, index)
            if fullwidth_match:
                key = fullwidth_match.group(1)
                _validate_reference_ids([key], ref_dict, fullwidth_match.group(0))
                output.append(f"[^{key}]")
                index = fullwidth_match.end()
                continue

        if "circled" in styles and text[index] in CIRCLED_NUM_MAP:
            key = _circled_to_number(text[index])
            _validate_reference_ids([key], ref_dict, text[index])
            output.append(f"[^{key}]")
            index += 1
            continue

        output.append(text[index])
        index += 1
    return "".join(output)


def convert_citations_in_text(text: str, ref_dict: dict[str, str]) -> str:
    """
    将文内的 [1] 引用标记转换为 pandoc 脚注 [^1]。

    策略：
    - 如果 ref_dict 中有对应编号，正文直接用 [^编号]
    - 连续多引用 [1][2][3] → [^1][^2][^3]
    - 范围引用 [1-3] → [^1],[^2],[^3]

    注意：不处理代码块内的内容。
    """
    if not ref_dict:
        return text
    return _replace_citations_in_text(text, ref_dict, {"bracket", "dot"})


def convert_ref_dict_to_footnotes(ref_dict: dict[str, str]) -> str:
    """
    将参考文献字典转为 pandoc footnote 定义字符串。
    格式：[^1]: 参考文献内容
    """
    lines = []
    for num in sorted(ref_dict.keys(), key=int):
        content = ref_dict[num].strip()
        lines.append(f"\n[^{num}]: {content}")
    return "\n".join(lines)


def is_in_code_block(lines: list[str], idx: int) -> bool:
    """兼容旧调用；使用单次扫描判断指定行是否位于围栏代码块。"""
    fence_char = None
    fence_length = 0
    for line_number, line in enumerate(lines):
        if line_number >= idx:
            break
        match = FENCE_RE.match(line)
        if not match:
            continue
        marker = match.group(1)
        if fence_char is None:
            fence_char, fence_length = marker[0], len(marker)
        elif marker[0] == fence_char and len(marker) >= fence_length:
            fence_char, fence_length = None, 0
    return fence_char is not None


def preprocess_markdown(input_path: str) -> str:
    """
    预处理 Markdown 文件：
    1. 分离正文和参考文献
    2. 正文引用 → 脚注标记
    3. 参考文献 → 脚注定义
    返回处理后的完整 Markdown 文本
    """
    with open(input_path, "r", encoding="utf-8") as f:
        content = f.read()

    lines = content.splitlines()
    body_lines, ref_dict, styles = parse_reference_section(lines)

    # 处理正文中的引用
    processed_body = []
    fence_char = None
    fence_length = 0
    for line in body_lines:
        fence_match = FENCE_RE.match(line)
        if fence_match:
            marker = fence_match.group(1)
            if fence_char is None:
                fence_char, fence_length = marker[0], len(marker)
            elif marker[0] == fence_char and len(marker) >= fence_length:
                fence_char, fence_length = None, 0
            processed_body.append(line)
            continue
        if fence_char is not None or not ref_dict:
            processed_body.append(line)
            continue

        processed_body.append(_replace_citations_in_text(line, ref_dict, styles))

    result = "\n".join(processed_body)

    # 追加脚注定义
    if ref_dict:
        footnote_text = convert_ref_dict_to_footnotes(ref_dict)
        result += "\n" + footnote_text + "\n"

    return result


def get_pandoc_cmd(input_md: str, output_docx: str, reference_doc: str | None = None) -> list[str]:
    """构建 pandoc 命令"""
    cmd = [
        "pandoc",
        input_md,
        "-o", output_docx,
        "--from", "markdown+footnotes+smart",
        "--to", "docx",
        "--lua-filter", str(STRIP_COLOR_FILTER),
        # 确保脚注正常工作
        "--wrap=none",
        # 文档元数据默认设置
        "--metadata", "lang=zh-CN",
    ]
    if reference_doc:
        cmd.extend(["--reference-doc", reference_doc])
    return cmd


def validate_docx(path: str) -> None:
    """验证输出是结构完整且包含正文的 DOCX。"""
    try:
        with zipfile.ZipFile(path) as archive:
            broken_member = archive.testzip()
            if broken_member:
                raise ConversionError(f"DOCX 压缩成员损坏: {broken_member}")
            if "word/document.xml" not in archive.namelist():
                raise ConversionError("DOCX 缺少 word/document.xml")
    except (OSError, zipfile.BadZipFile) as exc:
        raise ConversionError(f"输出不是有效 DOCX: {exc}") from exc


def count_docx_footnotes(path: str) -> int:
    """统计实际脚注，不包含 Word 内置的分隔符脚注。"""
    try:
        from xml.etree import ElementTree as ET

        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        footnote_tag = f"{{{namespace}}}footnote"
        id_attr = f"{{{namespace}}}id"
        with zipfile.ZipFile(path) as archive:
            if "word/footnotes.xml" not in archive.namelist():
                return 0
            root = ET.fromstring(archive.read("word/footnotes.xml"))
        return sum(
            1 for footnote in root.findall(footnote_tag)
            if int(footnote.get(id_attr, "-1")) > 0
        )
    except (ValueError, ET.ParseError, OSError, zipfile.BadZipFile) as exc:
        raise ConversionError(f"无法读取脚注结构: {exc}") from exc


def run_conversion_step(cmd: list[str], step_name: str) -> subprocess.CompletedProcess:
    """运行外部转换步骤，失败时抛出可供 CLI 统一处理的异常。"""
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
    except OSError as exc:
        raise ConversionError(f"{step_name}无法启动: {exc}") from exc
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "未知错误").strip()
        raise ConversionError(f"{step_name}失败: {detail}")
    return result


def new_temp_docx(directory: str) -> str:
    with tempfile.NamedTemporaryFile(suffix=".docx", dir=directory,
                                     delete=False) as tmp:
        return tmp.name


def convert(input_path: str, output_path: str | None = None,
            reference_doc: str | None = None, add_toc: bool = False,
            format_style: str = "auto") -> str:
    """
    主转换函数。

    参数：
        input_path: 输入的 .md 文件路径
        output_path: 输出的 .docx 路径（默认：同名 .docx）
        reference_doc: 参考样式模板 .docx（可选）
        format_style: 排版风格 - "faxue"(中国法学) / "sheke"(中国社会科学) / "auto"(自动检测)

    返回：
        输出文件路径
    """
    input_abs = os.path.abspath(input_path)
    if not os.path.exists(input_abs):
        raise FileNotFoundError(f"输入文件不存在: {input_abs}")
    if reference_doc:
        reference_doc = os.path.abspath(reference_doc)
        if not os.path.exists(reference_doc):
            raise FileNotFoundError(f"参考样式模板不存在: {reference_doc}")
    if format_style not in ("auto", "faxue", "sheke"):
        raise ConversionError(f"不支持的排版风格: {format_style}")

    if output_path is None:
        output_path = os.path.splitext(input_abs)[0] + ".docx"
    output_abs = os.path.abspath(output_path)
    output_dir = os.path.dirname(output_abs)
    if not os.path.isdir(output_dir):
        raise FileNotFoundError(f"输出目录不存在: {output_dir}")

    print(f"📄 输入: {input_abs}")
    print(f"📝 输出: {output_abs}")

    # 预处理 Markdown
    print("🔄 预处理：转换引用标记 → pandoc 脚注...")
    processed_md = preprocess_markdown(input_abs)

    # Markdown 和各阶段 DOCX 均写入临时文件，全部成功后再原子替换目标。
    with tempfile.NamedTemporaryFile(mode="w", suffix=".md",
                                      encoding="utf-8", delete=False) as tmp:
        tmp.write(processed_md)
        tmp_path = tmp.name
    working_docx = new_temp_docx(output_dir)
    temporary_docx_files = [working_docx]
    try:
        cmd = get_pandoc_cmd(tmp_path, working_docx, reference_doc)
        print(f"🔧 执行: {' '.join(cmd)}")
        result = run_conversion_step(cmd, "Pandoc 转换")
        if result.stderr:
            print(f"⚠️  Pandoc 警告:\n{result.stderr}")

        # ── 后处理：修复 Pandoc 生成的标题颜色残留 ──
        if os.path.exists(str(FIX_HEADING_SCRIPT)):
            print("🎨 后处理：修复标题颜色/样式残留...")
            heading_output = new_temp_docx(output_dir)
            temporary_docx_files.append(heading_output)
            result2 = run_conversion_step(
                [sys.executable, str(FIX_HEADING_SCRIPT), working_docx,
                 "--output", heading_output],
                "标题样式修复",
            )
            validate_docx(heading_output)
            os.replace(heading_output, working_docx)
            print(f"   {result2.stdout.strip()}")
        else:
            print("ℹ️  跳过标题颜色后处理（可选脚本未安装）")

        # ── 自动检测排版风格 ──
        if format_style == "auto":
            with open(input_abs, "r", encoding="utf-8") as source:
                raw_content = source.read()
            has_circled = CIRCLED_NUM_RE.search(raw_content) is not None
            has_fullwidth = FULLWIDTH_BRACKET_RE.search(raw_content) is not None
            has_sheke_abstract = bool(re.search(r"摘\s*要\s*[：:]", raw_content))
            if has_circled or has_fullwidth or has_sheke_abstract:
                format_style = "sheke"
                print("🔍 自动检测：《中国社会科学》格式")
            else:
                format_style = "faxue"
                print("🔍 自动检测：《中国法学》格式")

        format_scripts = {
            "faxue": SCRIPT_DIR / "zhongguo-faxue-format.py",
            "sheke": SCRIPT_DIR / "zhongguo-sheke-format.py",
        }
        format_names = {
            "faxue": "《中国法学》",
            "sheke": "《中国社会科学》",
        }
        fmt_script = format_scripts[format_style]
        fmt_name = format_names[format_style]
        if not os.path.exists(str(fmt_script)):
            raise ConversionError(f"缺少期刊格式化脚本: {fmt_script}")
        if add_toc and format_style == "sheke":
            raise ConversionError("《中国社会科学》格式暂不支持 --toc")

        print(f"📐 后处理：应用{fmt_name}格式...")
        formatted_output = new_temp_docx(output_dir)
        temporary_docx_files.append(formatted_output)
        format_cmd = [sys.executable, str(fmt_script), working_docx,
                      formatted_output]
        if add_toc:
            format_cmd.append("--toc")
        result3 = run_conversion_step(format_cmd, f"{fmt_name}格式化")
        validate_docx(formatted_output)
        os.replace(formatted_output, working_docx)
        print(f"   {result3.stdout.strip()}")

        validate_docx(working_docx)
        os.replace(working_docx, output_abs)
        size_kb = os.path.getsize(output_abs) / 1024
        footnote_count = count_docx_footnotes(output_abs)
        print(
            f"✅ 转换成功！输出文件: {output_abs} "
            f"({size_kb:.1f} KB，脚注 {footnote_count} 条)")
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        for temporary_path in temporary_docx_files:
            if os.path.exists(temporary_path):
                os.unlink(temporary_path)

    return output_abs


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Markdown → Word 转换器（法律写作专用）\n"
                    "自动将 [1][2] 引用标记转为 Word 脚注，输出纯黑文字。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python3 md2docx.py 论文.md
  python3 md2docx.py 论文.md 输出.docx
  python3 md2docx.py 论文.md --reference-doc 模板.docx
        """
    )
    parser.add_argument("input", help="输入 Markdown 文件 (.md)")
    parser.add_argument("output", nargs="?", default=None,
                        help="输出 Word 文件 (.docx)，默认同名")
    parser.add_argument("--reference-doc", "-r", default=None,
                        help="参考样式模板 .docx（可选）")
    parser.add_argument("--dry-run", "-n", action="store_true",
                        help="仅显示预处理结果，不实际转换")
    parser.add_argument("--toc", action="store_true",
                        help="在标题前插入静态目次（仅中国法学格式）")
    parser.add_argument("--format", "-f", choices=["faxue", "sheke", "auto"],
                        default="auto",
                        help="排版风格: faxue=中国法学, sheke=中国社会科学, auto=自动检测(默认)")

    args = parser.parse_args()

    try:
        if args.dry_run:
            processed = preprocess_markdown(args.input)
            print(processed)
        else:
            convert(args.input, args.output, args.reference_doc,
                    add_toc=args.toc, format_style=args.format)
    except (Md2DocxError, OSError) as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
