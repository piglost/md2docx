# Changelog

## v0.3.1 - 2026-06-21

- Treat an unnumbered first-section `引言` / `引  言` heading as a first-level heading across all supported formats, while preserving the original unnumbered title text.
- Display generated Word footnote references and footnote markers with square brackets, e.g. `[1]`.

## v0.3.0 - 2026-06-21

- Add a manually selected course-paper format (`-f course`) based on the extracted OOXML layout of《论公共数据开放的数据价值实现原则.docx》:
  - A4 page, 72pt top/bottom margins, and 90pt left/right margins,
  - Songti SC 12pt body text with 24pt first-line indent and 1.5x line spacing,
  - Heiti SC Light title and Heiti SC Medium headings,
  - Songti SC bold labels with 仿宋 abstract/keyword content,
  - 宋体 9pt footnotes with zero before/after spacing.
- Keep `auto` detection unchanged so existing journal workflows still choose only《中国法学》or《中国社会科学》unless `-f course` is specified.
- Add regression tests for the course-paper page geometry, fonts, paragraph spacing, heading hierarchy, and footnote typography.
- Update README and skill instructions with the new format option.

## v0.2.1 - 2026-06-21

- Refine《中国法学》formatting against the measured PDF layout of《深度伪造技术全链式刑事治理模式研究》:
  - set body text to Fangzheng Shusong at the closest Word size for the measured 10.8pt and fixed 16.5pt line spacing,
  - use mirrored margins, measured first-line indent, and measured abstract side indents,
  - update first-level headings to Heiti 14pt and second-level headings to Heiti 10.8pt,
  - use Fangzheng Fangsong page headers with a 0.75pt header rule,
  - use Fangzheng Shusong footnotes at the closest Word size for the measured 7.8pt and fixed 12pt line spacing.
- Update README and skill instructions with the new measured《中国法学》parameters.
- Add regression assertions for the measured page geometry, header typography, body spacing, abstract indents, heading sizes, and footnote typography.

## v0.2.0 - 2026-06-17

- Add regression tests for citation preprocessing, DOCX structure, footnotes, and journal-specific formatting.
- Support robust conversion failure handling so failed formatting does not overwrite output files.
- Improve citation parsing for numeric, fullwidth, circled, ranged, repeated, and continued references.
- Preserve bold, italic, underline, links, code spans, fenced code, and ordered list numbering during conversion.
- Update《中国法学》formatting to match the 2026年第3期 reference layout:
  - custom page size and margins,
  - Fangzheng title/body fonts,
  - fixed line spacing,
  - first-page handling, odd/even headers, and outside page numbers,
  - static TOC box with first-level headings only.
- Update《中国社会科学》formatting to match the 2026年第2期 reference layout:
  - `摘  要` label spacing,
  - Fangzheng title/author/abstract/body fonts,
  - mirror margins,
  - odd/even headers and outside page numbers,
  - circled footnote numbering and zero before/after footnote spacing.
- Refresh README and skill instructions with the current journal layout behavior.
