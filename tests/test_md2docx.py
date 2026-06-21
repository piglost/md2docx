from __future__ import annotations

import contextlib
import importlib.util
import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = lambda tag: f"{{{W_NS}}}{tag}"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = lambda tag: f"{{{R_NS}}}{tag}"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


md2docx = load_module("md2docx_under_test", ROOT / "md2docx.py")


def preprocess(markdown: str) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        source = Path(tmp) / "input.md"
        source.write_text(markdown, encoding="utf-8")
        return md2docx.preprocess_markdown(str(source))


def convert(markdown: str, style: str = "faxue", toc: bool = False):
    holder = tempfile.TemporaryDirectory()
    root = Path(holder.name)
    source = root / "input.md"
    output = root / "output.docx"
    source.write_text(markdown, encoding="utf-8")

    original_fixer = md2docx.FIX_HEADING_SCRIPT
    md2docx.FIX_HEADING_SCRIPT = root / "missing-heading-fixer.py"
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            md2docx.convert(str(source), str(output), format_style=style, add_toc=toc)
    finally:
        md2docx.FIX_HEADING_SCRIPT = original_fixer
    return output, holder


def read_xml(docx: Path, member: str) -> ET.Element:
    with zipfile.ZipFile(docx) as archive:
        return ET.fromstring(archive.read(member))


def paragraph_for_text(root: ET.Element, expected: str) -> ET.Element:
    for paragraph in root.iter(W("p")):
        text = "".join(node.text or "" for node in paragraph.iter(W("t")))
        if expected in text:
            return paragraph
    raise AssertionError(f"Paragraph containing {expected!r} not found")


def run_for_text(paragraph: ET.Element, expected: str) -> ET.Element:
    for run in paragraph.iter(W("r")):
        text = "".join(node.text or "" for node in run.iter(W("t")))
        if expected in text:
            return run
    raise AssertionError(f"Run containing {expected!r} not found")


class PreprocessTests(unittest.TestCase):
    def test_basic_citations_and_range_are_expanded(self):
        result = preprocess(
            "正文[1][2-3]\n\n## 参考文献\n"
            "1. 文献一\n2. 文献二\n3. 文献三\n"
        )
        self.assertIn("正文[^1][^2][^3]", result)
        self.assertIn("[^1]: 文献一", result)
        self.assertIn("[^3]: 文献三", result)

    def test_citation_replacement_preserves_following_space(self):
        result = preprocess("Text [1] next.\n\n## 参考文献\n1. Source\n")
        self.assertIn("Text [^1] next.", result)

    def test_inline_code_is_not_treated_as_citation(self):
        result = preprocess(
            "数组下标 `items[1]`，引用为 [1]。\n\n## 参考文献\n1. 文献一\n"
        )
        self.assertIn("`items[1]`", result)
        self.assertIn("引用为 [^1]", result)

    def test_fullwidth_and_circled_markers_in_code_or_links_are_untouched(self):
        result = preprocess(
            "代码 `items〔1〕⑪`，链接 [说明⑪](https://example.com)，"
            "引用〔1〕和⑪。\n\n## 注释\n"
            "〔1〕第一条注释。\n⑪ 第十一条注释。\n"
        )
        self.assertIn("`items〔1〕⑪`", result)
        self.assertIn("[说明⑪](https://example.com)", result)
        self.assertIn("引用[^1]和[^11]", result)

    def test_backtick_and_tilde_fenced_code_are_untouched(self):
        result = preprocess(
            "```python\na[1]\n```\n~~~python\nb[1]\n~~~\n"
            "正文[1]\n\n## 参考文献\n1. 文献一\n"
        )
        self.assertIn("```python\na[1]\n```", result)
        self.assertIn("~~~python\nb[1]\n~~~", result)

    def test_numeric_markdown_link_is_not_treated_as_citation(self):
        result = preprocess(
            "链接：[1](https://example.com)，图片：![1](image.png)，引用：[2]。\n\n"
            "## 参考文献\n1. 链接说明\n2. 文献二\n"
        )
        self.assertIn("[1](https://example.com)", result)
        self.assertIn("![1](image.png)", result)
        self.assertIn("引用：[^2]", result)

    def test_no_reference_section_leaves_numeric_brackets_unchanged(self):
        result = preprocess("数组下标 items[1]，没有参考文献章节。\n")
        self.assertIn("items[1]", result)
        self.assertNotIn("[^1]", result)

    def test_missing_reference_id_fails(self):
        with self.assertRaisesRegex(md2docx.CitationError, "2"):
            preprocess("正文引用[2]。\n\n## 参考文献\n1. 文献一\n")

    def test_descending_range_fails(self):
        with self.assertRaisesRegex(md2docx.CitationError, "3-1"):
            preprocess(
                "正文[3-1]。\n\n## 参考文献\n"
                "1. 文献一\n2. 文献二\n3. 文献三\n"
            )

    def test_excessive_range_fails(self):
        with self.assertRaisesRegex(md2docx.CitationError, "范围"):
            preprocess("正文[1-1000]。\n\n## 参考文献\n1. 文献一\n")

    def test_duplicate_reference_definition_fails(self):
        with self.assertRaisesRegex(md2docx.CitationError, "重复"):
            preprocess("正文[1]。\n\n## 参考文献\n1. 文献一\n1. 文献二\n")

    def test_circled_reference_11_is_supported(self):
        result = preprocess("正文⑪。\n\n## 注释\n⑪ 第十一条注释。\n")
        self.assertIn("正文[^11]", result)
        self.assertIn("[^11]: 第十一条注释。", result)

    def test_fullwidth_reference_continuation_is_preserved(self):
        result = preprocess(
            "正文〔1〕。\n\n## 参考文献\n"
            "〔1〕第一行。\n    第二行。\n"
        )
        self.assertIn("[^1]: 第一行。 第二行。", result)


class DocxStructureTests(unittest.TestCase):
    def test_generated_docx_is_valid_and_has_footnotes(self):
        output, holder = convert(
            "# 标题\n\n正文[1]。\n\n## 参考文献\n1. 文献一\n"
        )
        try:
            with zipfile.ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
                self.assertIn("word/document.xml", archive.namelist())
                self.assertIn("word/footnotes.xml", archive.namelist())
        finally:
            holder.cleanup()

    def test_body_bold_and_italic_are_preserved(self):
        output, holder = convert("# 标题\n\n普通 **加粗** 和 *斜体*。\n")
        try:
            root = read_xml(output, "word/document.xml")
            paragraph = paragraph_for_text(root, "普通 加粗 和 斜体")
            bold_run = run_for_text(paragraph, "加粗")
            italic_run = run_for_text(paragraph, "斜体")
            self.assertIsNotNone(bold_run.find(f"{W('rPr')}/{W('b')}"))
            self.assertIsNotNone(italic_run.find(f"{W('rPr')}/{W('i')}"))
        finally:
            holder.cleanup()

    def test_ordered_list_numbering_is_preserved(self):
        output, holder = convert("# 标题\n\n1. 第一项\n2. 第二项\n")
        try:
            root = read_xml(output, "word/document.xml")
            first_item = paragraph_for_text(root, "第一项")
            self.assertIsNotNone(first_item.find(f"{W('pPr')}/{W('numPr')}"))
        finally:
            holder.cleanup()

    def test_footnote_marker_keeps_reference_style(self):
        output, holder = convert(
            "# 标题\n\n正文[1]。\n\n## 参考文献\n1. 文献一\n"
        )
        try:
            root = read_xml(output, "word/footnotes.xml")
            marker_runs = [
                run for run in root.iter(W("r"))
                if run.find(W("footnoteRef")) is not None
            ]
            self.assertTrue(marker_runs)
            for run in marker_runs:
                style = run.find(f"{W('rPr')}/{W('rStyle')}")
                self.assertIsNotNone(style)
                self.assertEqual(style.get(W("val")), "FootnoteReference")
        finally:
            holder.cleanup()

    def test_faxue_matches_reference_page_geometry_and_headers(self):
        output, holder = convert(
            "# 《生态环境法典》的概念体系置换与话语体系建构\n\n"
            "吕忠梅\n\n内容提要 摘要正文。\n\n关键词 法典编纂\n\n"
            "## 一、一级标题\n\n正文。\n",
            style="faxue",
        )
        try:
            root = read_xml(output, "word/document.xml")
            section = root.find(f".//{W('sectPr')}")
            page_size = section.find(W("pgSz"))
            margins = section.find(W("pgMar"))
            self.assertEqual(page_size.get(W("w")), "10828")
            self.assertEqual(page_size.get(W("h")), "15080")
            self.assertEqual(margins.get(W("left")), "1304")
            self.assertEqual(margins.get(W("right")), "1361")

            settings = read_xml(output, "word/settings.xml")
            self.assertIsNotNone(settings.find(W("evenAndOddHeaders")))
            self.assertIsNotNone(settings.find(W("mirrorMargins")))

            with zipfile.ZipFile(output) as archive:
                self.assertIn("word/header1.xml", archive.namelist())
                self.assertIn("word/header2.xml", archive.namelist())
                odd_header = ET.fromstring(archive.read("word/header1.xml"))
                even_header = ET.fromstring(archive.read("word/header2.xml"))
                self.assertIn(
                    "《生态环境法典》的概念体系置换与话语体系建构",
                    "".join(node.text or "" for node in odd_header.iter(W("t"))),
                )
                self.assertIn(
                    "中国法学",
                    "".join(node.text or "" for node in even_header.iter(W("t"))),
                )
                header_border = odd_header.find(
                    f".//{W('pPr')}/{W('pBdr')}/{W('bottom')}")
                self.assertIsNotNone(header_border)
                self.assertEqual(header_border.get(W("sz")), "6")
                odd_header_run = run_for_text(
                    next(odd_header.iter(W("p"))),
                    "《生态环境法典》的概念体系置换与话语体系建构",
                )
                self.assertEqual(
                    odd_header_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                    "方正仿宋简体",
                )
                self.assertEqual(
                    odd_header_run.find(f"{W('rPr')}/{W('sz')}").get(W("val")),
                    "19",
                )
                self.assertIn(b"PAGE", archive.read("word/footer1.xml"))
                self.assertNotIn(b"ns0:Relationships", archive.read("word/_rels/document.xml.rels"))
        finally:
            holder.cleanup()

    def test_faxue_uses_reference_fonts_and_spacing(self):
        output, holder = convert(
            "# 《生态环境法典》的概念体系置换与话语体系建构\n\n"
            "吕忠梅\n\n内容提要 摘要正文。\n\n关键词 法典编纂\n\n"
            "## 一、一级标题\n\n正文第一段。\n\n### （一）二级标题\n\n正文第二段。\n",
            style="faxue",
        )
        try:
            root = read_xml(output, "word/document.xml")

            def style_for(text):
                paragraph = paragraph_for_text(root, text)
                run = run_for_text(paragraph, text)
                rpr = run.find(W("rPr"))
                fonts = rpr.find(W("rFonts"))
                size = rpr.find(W("sz"))
                spacing = paragraph.find(f"{W('pPr')}/{W('spacing')}")
                return paragraph, fonts, size, rpr, spacing

            _, fonts, size, rpr, _ = style_for("《生态环境法典》")
            self.assertEqual(fonts.get(W("eastAsia")), "方正小标宋简体")
            self.assertEqual(size.get(W("val")), "44")
            self.assertIsNone(rpr.find(W("b")))

            _, fonts, size, _, _ = style_for("吕忠梅")
            self.assertEqual(fonts.get(W("eastAsia")), "方正楷体简体")
            self.assertEqual(size.get(W("val")), "28")

            abstract = paragraph_for_text(root, "内容提要 摘要正文")
            label_run = run_for_text(abstract, "内容提要")
            content_run = run_for_text(abstract, "摘要正文")
            self.assertEqual(
                label_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "黑体",
            )
            self.assertEqual(
                content_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "方正楷体简体",
            )
            abstract_spacing = abstract.find(f"{W('pPr')}/{W('spacing')}")
            self.assertEqual(abstract_spacing.get(W("line")), "360")
            self.assertEqual(abstract_spacing.get(W("lineRule")), "exact")
            abstract_indent = abstract.find(f"{W('pPr')}/{W('ind')}")
            self.assertEqual(abstract_indent.get(W("left")), "454")
            self.assertEqual(abstract_indent.get(W("right")), "510")
            self.assertEqual(abstract_indent.get(W("firstLine")), "454")

            body, fonts, size, _, body_spacing = style_for("正文第一段")
            self.assertEqual(fonts.get(W("eastAsia")), "方正书宋简体")
            self.assertEqual(size.get(W("val")), "22")
            self.assertEqual(body_spacing.get(W("line")), "330")
            self.assertEqual(body_spacing.get(W("lineRule")), "exact")
            self.assertEqual(
                body.find(f"{W('pPr')}/{W('ind')}").get(W("firstLine")), "454")

            h1 = paragraph_for_text(root, "一、一级标题")
            h1_run = run_for_text(h1, "一、一级标题")
            self.assertEqual(
                h1_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "黑体",
            )
            self.assertEqual(h1_run.find(f"{W('rPr')}/{W('sz')}").get(W("val")), "28")
        finally:
            holder.cleanup()

    def test_faxue_toc_uses_outer_box_and_h1_only(self):
        output, holder = convert(
            "# 标题\n\n作者\n\n"
            "## 一、一级标题\n\n正文。\n\n### （一）二级标题\n\n正文。\n\n"
            "## 二、第二个一级标题\n\n正文。\n",
            style="faxue",
            toc=True,
        )
        try:
            root = read_xml(output, "word/document.xml")
            texts = [
                "".join(node.text or "" for node in paragraph.iter(W("t")))
                for paragraph in root.iter(W("p"))
            ]
            self.assertIn("目  次", texts)
            self.assertIn("一、一级标题", texts)
            self.assertIn("二、第二个一级标题", texts)
            toc_index = texts.index("目  次")
            toc_slice = texts[toc_index:toc_index + 4]
            self.assertNotIn("（一）二级标题", toc_slice)

            title_p = paragraph_for_text(root, "目  次")
            first_entry = paragraph_for_text(root, "一、一级标题")
            last_entry = paragraph_for_text(root, "二、第二个一级标题")
            self.assertIsNotNone(title_p.find(f"{W('pPr')}/{W('pBdr')}/{W('top')}"))
            self.assertIsNone(first_entry.find(f"{W('pPr')}/{W('pBdr')}/{W('bottom')}"))
            self.assertIsNotNone(last_entry.find(f"{W('pPr')}/{W('pBdr')}/{W('bottom')}"))
        finally:
            holder.cleanup()

    def test_sheke_footnote_position_uses_ooxml_child_element(self):
        output, holder = convert(
            "# 标题\n\n摘 要：测试。正文〔1〕。\n\n"
            "## 参考文献\n〔1〕文献一\n",
            style="sheke",
        )
        try:
            root = read_xml(output, "word/settings.xml")
            footnote_pr = root.find(W("footnotePr"))
            self.assertIsNotNone(footnote_pr)
            position = footnote_pr.find(W("pos"))
            self.assertIsNotNone(position)
            self.assertEqual(position.get(W("val")), "pageBottom")
            self.assertNotIn(W("pos"), footnote_pr.attrib)
        finally:
            holder.cleanup()

    def test_sheke_abstract_label_has_required_spacing(self):
        output, holder = convert("# 标题\n\n## 摘要\n\n摘要正文。\n", style="sheke")
        try:
            root = read_xml(output, "word/document.xml")
            texts = [
                "".join(node.text or "" for node in paragraph.iter(W("t")))
                for paragraph in root.iter(W("p"))
            ]
            self.assertIn("摘  要", texts)
            self.assertNotIn("摘要", texts)
        finally:
            holder.cleanup()

    def test_sheke_matches_reference_page_geometry_and_headers(self):
        output, holder = convert(
            "# 债权实现的程序法运行机理\n\n段文波\n\n"
            "摘 要：摘要正文。\n\n关键词：债权实现\n\n"
            "作者段文波，西南政法大学教授（重庆 401120）。\n\n正文。\n",
            style="sheke",
        )
        try:
            root = read_xml(output, "word/document.xml")
            section = root.find(f".//{W('sectPr')}")
            page_size = section.find(W("pgSz"))
            margins = section.find(W("pgMar"))
            self.assertEqual(page_size.get(W("w")), "11396")
            self.assertEqual(page_size.get(W("h")), "15874")
            self.assertEqual(margins.get(W("top")), "1440")
            self.assertEqual(margins.get(W("left")), "1538")
            self.assertEqual(margins.get(W("right")), "1090")

            settings = read_xml(output, "word/settings.xml")
            self.assertIsNotNone(settings.find(W("evenAndOddHeaders")))
            self.assertIsNotNone(settings.find(W("mirrorMargins")))

            with zipfile.ZipFile(output) as archive:
                names = archive.namelist()
                self.assertIn("word/header1.xml", names)
                self.assertIn("word/header2.xml", names)
                self.assertIn("word/header3.xml", names)
                self.assertIn("word/footer1.xml", names)
                self.assertIn("word/footer2.xml", names)
                relationships_xml = archive.read("word/_rels/document.xml.rels")
                content_types_xml = archive.read("[Content_Types].xml")
                self.assertNotIn(b"ns0:Relationships", relationships_xml)
                self.assertNotIn(b"ns0:Types", content_types_xml)
                odd_header = ET.fromstring(archive.read("word/header1.xml"))
                even_header = ET.fromstring(archive.read("word/header2.xml"))
                first_header = ET.fromstring(archive.read("word/header3.xml"))
                self.assertIn(
                    "债权实现的程序法运行机理",
                    "".join(node.text or "" for node in odd_header.iter(W("t"))),
                )
                self.assertIn(
                    "中国社会科学",
                    "".join(node.text or "" for node in even_header.iter(W("t"))),
                )
                self.assertEqual(
                    "".join(node.text or "" for node in first_header.iter(W("t"))),
                    "",
                )
                footer_xml = archive.read("word/footer1.xml")
                self.assertIn(b"PAGE", footer_xml)
        finally:
            holder.cleanup()

    def test_sheke_uses_reference_fonts_sizes_and_exact_line_spacing(self):
        output, holder = convert(
            "# 债权实现的程序法运行机理\n\n段文波\n\n"
            "摘 要：摘要正文。\n\n关键词：债权实现\n\n"
            "作者段文波，西南政法大学教授（重庆 401120）。\n\n"
            "正文第一段。\n\n## 一、一级标题\n\n正文第二段。\n",
            style="sheke",
        )
        try:
            root = read_xml(output, "word/document.xml")

            def style_for(text):
                paragraph = paragraph_for_text(root, text)
                run = run_for_text(paragraph, text)
                rpr = run.find(W("rPr"))
                fonts = rpr.find(W("rFonts"))
                size = rpr.find(W("sz"))
                spacing = paragraph.find(f"{W('pPr')}/{W('spacing')}")
                return paragraph, fonts, size, rpr, spacing

            _, fonts, size, rpr, _ = style_for("债权实现的程序法运行机理")
            self.assertEqual(fonts.get(W("eastAsia")), "方正小标宋简体")
            self.assertEqual(size.get(W("val")), "52")
            self.assertIsNone(rpr.find(W("b")))

            _, fonts, size, rpr, _ = style_for("段文波")
            self.assertEqual(fonts.get(W("eastAsia")), "方正楷体简体")
            self.assertEqual(size.get(W("val")), "30")
            self.assertEqual(rpr.find(W("spacing")).get(W("val")), "180")

            abstract = paragraph_for_text(root, "摘  要：摘要正文")
            label_run = run_for_text(abstract, "摘  要：")
            content_run = run_for_text(abstract, "摘要正文")
            self.assertEqual(
                label_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "方正黑体简体",
            )
            self.assertEqual(
                content_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "方正仿宋简体",
            )
            self.assertEqual(
                label_run.find(f"{W('rPr')}/{W('sz')}").get(W("val")), "20")
            abstract_spacing = abstract.find(f"{W('pPr')}/{W('spacing')}")
            self.assertEqual(abstract_spacing.get(W("line")), "338")
            self.assertEqual(abstract_spacing.get(W("lineRule")), "exact")
            abstract_indent = abstract.find(f"{W('pPr')}/{W('ind')}")
            self.assertEqual(abstract_indent.get(W("left")), "600")
            self.assertEqual(abstract_indent.get(W("right")), "600")

            body, fonts, size, _, body_spacing = style_for("正文第一段")
            self.assertEqual(fonts.get(W("eastAsia")), "方正书宋简体")
            self.assertEqual(size.get(W("val")), "22")
            self.assertEqual(body_spacing.get(W("line")), "372")
            self.assertEqual(body_spacing.get(W("lineRule")), "exact")
            self.assertEqual(
                body.find(f"{W('pPr')}/{W('ind')}").get(W("firstLine")), "480")

            author_unit = paragraph_for_text(root, "作者段文波")
            border = author_unit.find(f"{W('pPr')}/{W('pBdr')}/{W('bottom')}")
            self.assertIsNotNone(border)
        finally:
            holder.cleanup()

    def test_sheke_footnotes_use_reference_size_and_circled_numbering(self):
        output, holder = convert(
            "# 标题\n\n摘 要：测试。正文〔1〕。\n\n"
            "## 参考文献\n〔1〕文献一\n",
            style="sheke",
        )
        try:
            footnotes = read_xml(output, "word/footnotes.xml")
            text_run = next(
                run for run in footnotes.iter(W("r"))
                if run.find(W("footnoteRef")) is None
                and "文献一" in "".join(t.text or "" for t in run.iter(W("t")))
            )
            rpr = text_run.find(W("rPr"))
            self.assertEqual(rpr.find(W("rFonts")).get(W("eastAsia")), "方正仿宋简体")
            self.assertEqual(rpr.find(W("sz")).get(W("val")), "20")

            settings = read_xml(output, "word/settings.xml")
            footnote_pr = settings.find(W("footnotePr"))
            self.assertEqual(
                footnote_pr.find(W("numFmt")).get(W("val")),
                "decimalEnclosedCircle",
            )
        finally:
            holder.cleanup()

    def test_course_paper_matches_extracted_document_format(self):
        output, holder = convert(
            "# 论公共数据开放的数据价值实现原则\n\n"
            "格式 Demo\n\n"
            "摘要：摘要正文。\n\n关键词：公共数据开放；数据价值\n\n"
            "## 一、问题的提出\n\n正文第一段。\n\n"
            "### （一）价值实现原则的规范定位\n\n正文第二段。[1]\n\n"
            "## 参考文献\n1. 文献一\n",
            style="course",
        )
        try:
            root = read_xml(output, "word/document.xml")
            section = root.find(f".//{W('sectPr')}")
            page_size = section.find(W("pgSz"))
            margins = section.find(W("pgMar"))
            self.assertEqual(page_size.get(W("w")), "11906")
            self.assertEqual(page_size.get(W("h")), "16838")
            self.assertEqual(margins.get(W("top")), "1440")
            self.assertEqual(margins.get(W("bottom")), "1440")
            self.assertEqual(margins.get(W("left")), "1800")
            self.assertEqual(margins.get(W("right")), "1800")

            def style_for(text):
                paragraph = paragraph_for_text(root, text)
                run = run_for_text(paragraph, text)
                rpr = run.find(W("rPr"))
                fonts = rpr.find(W("rFonts"))
                size = rpr.find(W("sz"))
                spacing = paragraph.find(f"{W('pPr')}/{W('spacing')}")
                indent = paragraph.find(f"{W('pPr')}/{W('ind')}")
                return paragraph, fonts, size, rpr, spacing, indent

            _, fonts, size, rpr, _, _ = style_for("论公共数据开放")
            self.assertEqual(fonts.get(W("eastAsia")), "Heiti SC Light")
            self.assertEqual(size.get(W("val")), "36")
            self.assertIsNotNone(rpr.find(W("b")))

            abstract = paragraph_for_text(root, "摘要： 摘要正文")
            label_run = run_for_text(abstract, "摘要：")
            content_run = run_for_text(abstract, "摘要正文")
            self.assertEqual(
                label_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "Songti SC",
            )
            self.assertIsNotNone(label_run.find(f"{W('rPr')}/{W('b')}"))
            self.assertEqual(
                content_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "仿宋",
            )
            abstract_indent = abstract.find(f"{W('pPr')}/{W('ind')}")
            self.assertEqual(abstract_indent.get(W("left")), "480")
            self.assertEqual(abstract_indent.get(W("right")), "480")
            self.assertEqual(abstract_indent.get(W("firstLine")), "420")

            body, fonts, size, _, spacing, indent = style_for("正文第一段")
            self.assertEqual(fonts.get(W("eastAsia")), "Songti SC")
            self.assertEqual(size.get(W("val")), "24")
            self.assertEqual(spacing.get(W("line")), "360")
            self.assertEqual(spacing.get(W("lineRule")), "auto")
            self.assertEqual(indent.get(W("firstLine")), "480")
            self.assertEqual(
                body.find(f"{W('pPr')}/{W('jc')}").get(W("val")),
                "both",
            )

            h1 = paragraph_for_text(root, "一、问题的提出")
            h1_run = run_for_text(h1, "一、问题的提出")
            self.assertEqual(
                h1_run.find(f"{W('rPr')}/{W('rFonts')}").get(W("eastAsia")),
                "Heiti SC Medium",
            )
            self.assertEqual(h1_run.find(f"{W('rPr')}/{W('sz')}").get(W("val")), "28")
            self.assertIsNotNone(h1_run.find(f"{W('rPr')}/{W('b')}"))

            footnotes = read_xml(output, "word/footnotes.xml")
            text_run = next(
                run for run in footnotes.iter(W("r"))
                if run.find(W("footnoteRef")) is None
                and "文献一" in "".join(t.text or "" for t in run.iter(W("t")))
            )
            rpr = text_run.find(W("rPr"))
            self.assertEqual(rpr.find(W("rFonts")).get(W("eastAsia")), "宋体")
            self.assertEqual(rpr.find(W("sz")).get(W("val")), "18")
        finally:
            holder.cleanup()

    def test_footnote_paragraph_spacing_is_zero_in_both_formats(self):
        markdown_by_style = {
            "faxue": "# 标题\n\n正文[1]。\n\n## 参考文献\n1. 文献一\n",
            "sheke": "# 标题\n\n摘 要：测试。正文〔1〕。\n\n## 参考文献\n〔1〕文献一\n",
            "course": "# 标题\n\n正文[1]。\n\n## 参考文献\n1. 文献一\n",
        }
        for style, markdown in markdown_by_style.items():
            with self.subTest(style=style):
                output, holder = convert(markdown, style=style)
                try:
                    root = read_xml(output, "word/footnotes.xml")
                    actual_footnotes = [
                        footnote for footnote in root.findall(W("footnote"))
                        if int(footnote.get(W("id"), "-1")) > 0
                    ]
                    self.assertTrue(actual_footnotes)
                    for footnote in actual_footnotes:
                        for paragraph in footnote.iter(W("p")):
                            spacing = paragraph.find(f"{W('pPr')}/{W('spacing')}")
                            self.assertIsNotNone(spacing)
                            self.assertEqual(spacing.get(W("before")), "0")
                            self.assertEqual(spacing.get(W("after")), "0")
                            if style == "faxue":
                                self.assertEqual(spacing.get(W("line")), "240")
                                text_runs = [
                                    run for run in paragraph.iter(W("r"))
                                    if run.find(W("footnoteRef")) is None
                                    and "".join(t.text or "" for t in run.iter(W("t"))).strip()
                                ]
                                self.assertTrue(text_runs)
                                rpr = text_runs[0].find(W("rPr"))
                                self.assertEqual(
                                    rpr.find(W("rFonts")).get(W("eastAsia")),
                                    "方正书宋简体",
                                )
                                self.assertEqual(rpr.find(W("sz")).get(W("val")), "16")
                            elif style == "course":
                                self.assertEqual(spacing.get(W("line")), "240")
                                text_runs = [
                                    run for run in paragraph.iter(W("r"))
                                    if run.find(W("footnoteRef")) is None
                                    and "".join(t.text or "" for t in run.iter(W("t"))).strip()
                                ]
                                self.assertTrue(text_runs)
                                rpr = text_runs[0].find(W("rPr"))
                                self.assertEqual(
                                    rpr.find(W("rFonts")).get(W("eastAsia")),
                                    "宋体",
                                )
                                self.assertEqual(rpr.find(W("sz")).get(W("val")), "18")
                finally:
                    holder.cleanup()

    def test_formatting_failure_makes_conversion_fail(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "input.md"
            output = root / "output.docx"
            fake_scripts = root / "scripts"
            fake_scripts.mkdir()
            source.write_text("# 标题\n\n正文。\n", encoding="utf-8")
            (fake_scripts / "zhongguo-faxue-format.py").write_text(
                "import sys\nprint('forced failure', file=sys.stderr)\nsys.exit(7)\n",
                encoding="utf-8",
            )

            original_dir = md2docx.SCRIPT_DIR
            original_fixer = md2docx.FIX_HEADING_SCRIPT
            md2docx.SCRIPT_DIR = fake_scripts
            md2docx.FIX_HEADING_SCRIPT = root / "missing-heading-fixer.py"
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    with self.assertRaisesRegex(md2docx.ConversionError, "forced failure"):
                        md2docx.convert(str(source), str(output), format_style="faxue")
                self.assertFalse(output.exists())
            finally:
                md2docx.SCRIPT_DIR = original_dir
                md2docx.FIX_HEADING_SCRIPT = original_fixer


if __name__ == "__main__":
    unittest.main(verbosity=2)
