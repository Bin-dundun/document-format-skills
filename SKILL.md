---
name: document-format-skills
description: Chinese official-document formatting toolkit for .docx/.doc/.wps workflows. Use when Codex needs to diagnose and format Chinese government-style documents, fix mixed Chinese/English punctuation and spacing, normalize tables and page numbers, preserve or output Word revision marks, or convert plain text or Markdown into an official-format DOCX.
---

# Document Format Skills

Use these scripts to clean and format Chinese Word documents from the command line. Prefer `scripts/process.py` for normal work because it mirrors the desktop app's core pipeline without the GUI.

## Quick Workflow

Run one smart pass when the user wants the document cleaned end to end:

```bash
uv run --with python-docx python scripts/process.py smart input.docx output.docx --preset official
```

Run diagnostics only:

```bash
uv run --with python-docx python scripts/process.py analyze input.docx
uv run --with python-docx python scripts/process.py analyze input.docx --json
```

Run only punctuation/spacing cleanup:

```bash
uv run --with python-docx python scripts/process.py punctuation input.docx output.docx --space-mode keep_en_boundary
```

Run only formatting:

```bash
uv run --with python-docx python scripts/process.py format input.docx output.docx --preset official
```

On Windows, `.doc` and `.wps` input/output are supported through WPS or Microsoft Word COM automation:

```bash
uv run --with python-docx --with pywin32 python scripts/process.py smart input.wps output.wps
```

## Scripts

| Script | Use |
| --- | --- |
| `scripts/process.py` | One-shot CLI for `smart`, `analyze`, `punctuation`, and `format`; handles `.doc/.wps` conversion on Windows. |
| `scripts/formatter.py` | Apply the official-document format, page numbers, table cleanup, revision marks, and macOS font fallback. |
| `scripts/punctuation.py` | Fix punctuation while preserving run formatting; supports spacing strategies. |
| `scripts/from_text.py` | Create a DOCX from `.txt` or Markdown, then optionally run smart formatting. |
| `scripts/analyzer.py` | Lower-level diagnostic script. |
| `scripts/converter.py` | Windows-only `.doc/.wps` conversion helpers. |

## Formatting Options

The only exposed format is `official`:

- Main title: 小二号（18 pt）方正小标宋简体, centered.
- Recipient, body, signature, date, attachment, and closing: 三号（16 pt）仿宋_GB2312; body uses two-character first-line indent and fixed 30 pt line spacing.
- Level-1 headings: 三号（16 pt）方正黑体_GBK; level-2 headings: 三号（16 pt）楷体_GB2312.
- Latin text: Times New Roman.
- Chinese double and single quotation marks use the standard Chinese Songti glyphs.
- Margins: top/bottom 25.4 mm; left/right 31.8 mm.
- Page numbers: enabled by default, with digits set in Times New Roman; style, position, offset, and replacement behavior are configurable.

Useful flags:

```bash
--revision
--deep-clean
--smart-table-align
--no-page-number
--page-number-style dash|plain|page_text|page_total
--page-number-position outside|left|center|right
--page-number-offset-mm 7
--no-bold-serial
```

## Punctuation And Spacing

Punctuation cleanup protects URLs, email addresses, Windows paths, time values like `9:30`, and standards like `ISO 9001:2015`. It fixes brackets, colons, semicolons, question/exclamation marks, Chinese comma/period contexts, ellipses, dashes, and paired quotes.

Spacing modes:

- `remove_all`: delete half-width and full-width spaces.
- `keep_en_boundary`: remove Chinese-to-Chinese spaces but keep exactly one space between Chinese and English/digits.
- `keep_all`: leave spaces unchanged.

## Text Or Markdown To DOCX

Generate and format a document from text:

```bash
uv run --with python-docx python scripts/from_text.py input.md output.docx --title "工作方案"
```

Markdown mode detects headings, bold spans, ordered/unordered lists, quotes, and fenced code blocks. `#` becomes the main title, `##` becomes `一、`, `###` becomes `（一）`, and deeper headings become numbered lower-level headings.

Use `--no-process` to only create the raw DOCX.

## Implementation Notes

- `.docx` processing needs only `python-docx`.
- `.doc/.wps` conversion needs Windows plus WPS Office or Microsoft Word and `pywin32`.
- Page number handling avoids overwriting non-page footer content and can replace existing page-number footers when requested.
- Default table formatting preserves original alignment; use `--smart-table-align` for numeric/right and short-text/center alignment.
- macOS font handling keeps installed official fonts when present and falls back to compatible system fonts only when detection confirms the original is missing.
