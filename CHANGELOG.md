# Changelog

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
