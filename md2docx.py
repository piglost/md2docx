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
import shutil
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
STRIP_COLOR_FILTER = SCRIPT_DIR / "strip-color.lua"

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

# 文内引用模式（不参考文末列表，直接全文转换）
# 匹配独立的 [数字] 或 [数字,数字] 或 [数字-数字]
INLINE_CITE_RE = re.compile(
    r"\[(\d+(?:[,，]\s*\d+)*(?:[-–—]\d+)?)\]"
)

# 匹配连续的多个引用：[1][2][3]
MULTI_CITE_RE = re.compile(
    r"(?:\[\d+(?:[,，]\s*\d+)*(?:[-–—]\d+)?\]\s*)+"
)

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
    boundary = find_circled_ref_boundary(lines)
    if boundary is None:
        return lines, {}

    body = lines[:boundary]
    ref_lines = lines[boundary + 1:]
    ref_dict = {}
    current_key = None
    current_text = []

    for line in ref_lines:
        stripped = line.strip()
        # 跳过子标题
        if REF_SUBHEADING_RE.match(stripped):
            continue
        # 匹配 ① 内容 格式
        m = re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*(.*)", stripped)
        if m:
            if current_key is not None and current_text:
                ref_dict[current_key] = " ".join(current_text).strip()
            # 从行首提取圆圈数字
            circled = stripped[0]
            idx = CIRCLED_NUM_MAP.index(circled) + 1
            current_key = str(idx)
            current_text = [m.group(1).strip()] if m.group(1).strip() else []
            continue
        # 续行
        if current_key is not None and stripped:
            current_text.append(stripped)

    if current_key is not None and current_text:
        ref_dict[current_key] = " ".join(current_text).strip()

    return body, ref_dict


def convert_circled_citations(text: str, ref_dict: dict[str, str]) -> str:
    """将 ①②③ 标记转为 pandoc 脚注 [^N]"""
    if not ref_dict:
        return text
    def replace(match):
        circled = match.group(0)
        idx = CIRCLED_NUM_MAP.index(circled) + 1
        key = str(idx)
        if key in ref_dict:
            return f"[^{key}]"
        return match.group(0)  # 无对应定义则保留原样
    return CIRCLED_NUM_RE.sub(replace, text)


def find_ref_section_boundary(lines: list[str]) -> int | None:
    """
    找到参考文献 section 的起始行索引。
    返回 None 表示没有专门的参考文献节。
    """
    for ref_title in REF_SECTION_TITLES:
        # Markdown heading: # 参考文献, ## 参考文献, ### 参考文献
        heading_pat = re.compile(rf"^#{{1,4}}\s+{re.escape(ref_title)}\s*$")
        for i, line in enumerate(lines):
            if heading_pat.match(line.strip()):
                return i
    return None


def split_document(lines: list[str]) -> tuple[list[str], dict[str, str]]:
    """
    将文档分成正文部分和参考文献部分。
    返回：(正文行列表, {编号: 参考文献内容})
    """
    ref_boundary = find_ref_section_boundary(lines)
    if ref_boundary is None:
        # 没有独立的参考文献节，全文视为正文
        return lines, {}

    body_lines = lines[:ref_boundary]
    ref_lines = lines[ref_boundary + 1:]  # 跳过标题行
    ref_dict = {}

    current_key = None
    current_text = []

    for line in ref_lines:
        stripped = line.strip()

        # 跳过子标题（如 ### 一、规范性文件）
        if REF_SUBHEADING_RE.match(stripped):
            continue

        # 尝试匹配 [1] 格式
        m = REF_ENTRY_BRACKET_RE.match(stripped)
        if m:
            if current_key is not None and current_text:
                ref_dict[current_key] = " ".join(current_text).strip()
            current_key = m.group(1)
            current_text = [m.group(2).strip()] if m.group(2) else []
            current_text = [x for x in current_text if x]
            continue

        # 尝试匹配 1. 格式
        m = REF_ENTRY_DOT_RE.match(stripped)
        if m:
            if current_key is not None and current_text:
                ref_dict[current_key] = " ".join(current_text).strip()
            current_key = m.group(1)
            content = m.group(2).strip()
            current_text = [content] if content else []
            continue

        # 空行：结束当前条目
        if not stripped:
            if current_key is not None and current_text:
                ref_dict[current_key] = " ".join(current_text).strip()
                current_key = None
                current_text = []
            continue

        # 续行
        if current_key is not None:
            current_text.append(stripped)

    # 保存最后一条
    if current_key is not None and current_text:
        ref_dict[current_key] = " ".join(current_text).strip()

    return body_lines, ref_dict


def convert_citations_in_text(text: str, ref_dict: dict[str, str]) -> str:
    """
    将文内的 [1] 引用标记转换为 pandoc 脚注 [^1]。

    策略：
    - 如果 ref_dict 中有对应编号，正文直接用 [^编号]
    - 连续多引用 [1][2][3] → [^1][^2][^3]
    - 范围引用 [1-3] → [^1],[^2],[^3]

    注意：不处理代码块内的内容。
    """
    def replace_multi_cite(match):
        """替换一组连续的引用标记"""
        original = match.group(0)
        # 提取所有数字
        nums = re.findall(r"\[(\d+(?:[,，]\s*\d+)*(?:[-–—]\d+)?)\]", original)
        expanded = []
        for n_str in nums:
            if "-" in n_str or "–" in n_str or "—" in n_str:
                # 范围：[1-3]
                parts = re.split(r"[-–—]", n_str)
                if len(parts) == 2:
                    try:
                        start, end = int(parts[0]), int(parts[1])
                        for k in range(start, end + 1):
                            expanded.append(str(k))
                    except ValueError:
                        expanded.append(n_str)
            elif "," in n_str or "，" in n_str:
                # 逗号分隔：[1,2,3]
                for part in re.split(r"[,，]\s*", n_str):
                    expanded.append(part.strip())
            else:
                expanded.append(n_str)

        # 去重保持顺序
        seen = set()
        uniq = []
        for x in expanded:
            if x not in seen:
                seen.add(x)
                uniq.append(x)

        return "".join(f"[^{n}]" for n in uniq)

    # 按优先级：先处理连续多引用，再处理单引用
    text = MULTI_CITE_RE.sub(replace_multi_cite, text)

    return text


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
    """检查某行是否在代码块内"""
    in_block = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("```"):
            if i < idx:
                in_block = not in_block
    return in_block


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
    body_lines, ref_dict = split_document(lines)

    # 检测是否有 〔1〕 格式的参考文献需要转为脚注
    if not ref_dict:
        fullwidth_refs = FULLWIDTH_BRACKET_RE.findall(content)
        if fullwidth_refs:
            new_ref_dict = {}
            new_body = []
            in_ref_section = False
            for line in lines:
                stripped = line.strip()
                if re.match(r"^#+\s+参考文献", stripped):
                    in_ref_section = True
                    continue
                if in_ref_section:
                    m = re.match(r"^〔(\d+)〕\s*(.*)", stripped)
                    if m:
                        new_ref_dict[m.group(1)] = m.group(2).strip()
                    continue
                new_body.append(line)
            if new_ref_dict:
                body_lines = []
                for line in new_body:
                    body_lines.append(FULLWIDTH_BRACKET_RE.sub(
                        lambda m: f"[^{m.group(1)}]", line))
                ref_dict = new_ref_dict

    # 检测正文是否有 ①②③ 标记，优先使用圆圈数字脚注
    has_circled = any(c in "①②③④⑤⑥⑦⑧⑨⑩" for line in lines for c in line)
    if has_circled:
        body_lines, ref_dict = split_circled_footnotes(lines)
        use_circled = bool(ref_dict)
    else:
        use_circled = False

    # 处理正文中的引用
    processed_body = []
    for i, line in enumerate(body_lines):
        if is_in_code_block(body_lines, i) or line.strip().startswith("``"):
            processed_body.append(line)
            continue
        if use_circled:
            processed_body.append(convert_circled_citations(line, ref_dict))
        else:
            processed_body.append(convert_citations_in_text(line, ref_dict))

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
    if reference_doc and os.path.exists(reference_doc):
        cmd.extend(["--reference-doc", reference_doc])
    return cmd


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

    if output_path is None:
        output_path = os.path.splitext(input_abs)[0] + ".docx"
    output_abs = os.path.abspath(output_path)

    print(f"📄 输入: {input_abs}")
    print(f"📝 输出: {output_abs}")

    # 预处理 Markdown
    print("🔄 预处理：转换引用标记 → pandoc 脚注...")
    processed_md = preprocess_markdown(input_abs)

    # 写到临时文件
    with tempfile.NamedTemporaryFile(mode="w", suffix=".md",
                                      encoding="utf-8", delete=False) as tmp:
        tmp.write(processed_md)
        tmp_path = tmp.name

    try:
        # 调用 pandoc
        cmd = get_pandoc_cmd(tmp_path, output_abs, reference_doc)
        print(f"🔧 执行: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"❌ Pandoc 错误:\n{result.stderr}")
            sys.exit(1)

        if result.stderr:
            print(f"⚠️  Pandoc 警告:\n{result.stderr}")

    finally:
        os.unlink(tmp_path)

    # ── 后处理：修复 Pandoc 生成的标题颜色残留 ──
    if os.path.exists(str(FIX_HEADING_SCRIPT)):
        print("🎨 后处理：修复标题颜色/样式残留...")
        import tempfile as tmpmod
        with tmpmod.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_out:
            tmp_out_path = tmp_out.name
        try:
            result2 = subprocess.run(
                [
                    sys.executable, str(FIX_HEADING_SCRIPT),
                    output_abs,
                    "--output", tmp_out_path,
                ],
                capture_output=True, text=True,
            )
            if result2.returncode == 0:
                shutil.move(tmp_out_path, output_abs)
                print(f"   {result2.stdout.strip()}")
            else:
                print(f"   ⚠️ 后处理跳过: {result2.stderr.strip()}")
                os.unlink(tmp_out_path)
        except Exception as e:
            print(f"   ⚠️ 后处理异常: {e}")
            if os.path.exists(tmp_out_path):
                os.unlink(tmp_out_path)
    else:
        print("ℹ️  跳过颜色后处理（脚本未安装）")

    # ── 自动检测排版风格 ──
    if format_style == "auto":
        with open(input_abs, "r", encoding="utf-8") as f:
            raw_content = f.read()
        has_circled = any(c in "①②③④⑤⑥⑦⑧⑨⑩" for c in raw_content)
        has_fullwidth = "〔" in raw_content and "〕" in raw_content
        has_sheke_abstract = bool(re.search(r"摘\s*要\s*[：:]", raw_content))
        if has_circled or has_fullwidth or has_sheke_abstract:
            format_style = "sheke"
            print("🔍 自动检测：《中国社会科学》格式")
        else:
            format_style = "faxue"
            print("🔍 自动检测：《中国法学》格式")

    # ── 后处理：应用期刊格式 ──
    format_scripts = {
        "faxue": SCRIPT_DIR / "zhongguo-faxue-format.py",
        "sheke": SCRIPT_DIR / "zhongguo-sheke-format.py",
    }
    format_names = {
        "faxue": "《中国法学》",
        "sheke": "《中国社会科学》",
    }
    fmt_script = format_scripts.get(format_style, format_scripts["faxue"])
    fmt_name = format_names.get(format_style, "默认")
    if os.path.exists(str(fmt_script)):
        print(f"📐 后处理：应用{fmt_name}格式...")
        import tempfile as tmpmod2
        with tmpmod2.NamedTemporaryFile(suffix=".docx", delete=False) as tmp2:
            zgf_tmp = tmp2.name
        try:
            zgf_cmd = [sys.executable, str(fmt_script), output_abs, zgf_tmp]
            if add_toc:
                zgf_cmd.append("--toc")
            result3 = subprocess.run(zgf_cmd, capture_output=True, text=True)
            if result3.returncode == 0:
                shutil.move(zgf_tmp, output_abs)
                print(f"   {result3.stdout.strip()}")
            else:
                print(f"   ⚠️ 格式应用失败: {result3.stderr.strip()}")
                if os.path.exists(zgf_tmp):
                    os.unlink(zgf_tmp)
        except Exception as e:
            print(f"   ⚠️ 格式异常: {e}")
            if os.path.exists(zgf_tmp):
                os.unlink(zgf_tmp)
    else:
        # 回退：仅统一宋体
        print("🔤 后处理：统一字体 → 宋体 + Times New Roman...")
        font_ok = normalize_font_to_songti(output_abs)
        if font_ok:
            print("   字体已统一设置")
        else:
            print("   ⚠️ 字体设置失败")

    # 验证输出
    if os.path.exists(output_abs):
        size_kb = os.path.getsize(output_abs) / 1024
        print(f"✅ 转换成功！输出文件: {output_abs} ({size_kb:.1f} KB)")
    else:
        print(f"❌ 输出文件未生成")
        sys.exit(1)

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
                        help="在标题前插入自动目次（Word 域代码）")
    parser.add_argument("--format", "-f", choices=["faxue", "sheke", "auto"],
                        default="auto",
                        help="排版风格: faxue=中国法学, sheke=中国社会科学, auto=自动检测(默认)")

    args = parser.parse_args()

    if args.dry_run:
        processed = preprocess_markdown(args.input)
        print(processed)
    else:
        convert(args.input, args.output, args.reference_doc,
                add_toc=args.toc, format_style=args.format)


if __name__ == "__main__":
    main()
