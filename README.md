# document-format-skills 公文格式助手

只面向中文公文排版的命令行 Skill，用于诊断 Word 文档格式问题、修复中英文标点和空格混用、统一公文格式、处理页码和表格，并从纯文本或 Markdown 生成规范 DOCX。

## 功能

- 智能一键处理：标点/空格清理 + 公文格式统一。
- 格式诊断：检查标点、序号、段落和字体问题。
- 提供统一的中文公文排版格式。
- 支持页码样式、位置、距版心偏移、替换已有页码，并避免覆盖非页码页脚内容。
- 支持表格格式统一、智能对齐和 Word 修订标记。
- Windows 下可借助 WPS Office 或 Microsoft Word 处理 `.doc` / `.wps`。
- 支持从纯文本或 Markdown 生成并格式化 DOCX。

## 公文格式要求

| 项目 | 要求 |
| --- | --- |
| 页面 | 上下页边距 25.4mm，左右页边距 31.8mm |
| 主标题 | 方正小标宋简体，小二号（18pt），居中 |
| 正文 | 仿宋_GB2312，三号（16pt），首行缩进2字符，固定行距30磅 |
| 一级标题 | 方正黑体_GBK，三号（16pt），首行缩进2字符 |
| 二级标题 | 楷体_GB2312，三号（16pt），首行缩进2字符 |
| 英文及数字 | Times New Roman |
| 行距 | 全文全选后固定值30磅 |
| 页码 | 默认添加，可设置样式、位置、偏移和是否替换已有页码 |

## 环境要求

- Python 3.8+
- `python-docx`
- Windows 下处理 `.doc/.wps` 需要 `pywin32`、WPS Office 或 Microsoft Word

推荐使用 `uv` 临时安装依赖：

```bash
uv run --with python-docx python scripts/process.py --help
```

## 快速开始

智能一键处理：

```bash
uv run --with python-docx python scripts/process.py smart input.docx output.docx --preset official
```

只做诊断：

```bash
uv run --with python-docx python scripts/process.py analyze input.docx
uv run --with python-docx python scripts/process.py analyze input.docx --json
```

只修复标点和空格：

```bash
uv run --with python-docx python scripts/process.py punctuation input.docx output.docx --space-mode keep_en_boundary
```

只应用公文格式：

```bash
uv run --with python-docx python scripts/process.py format input.docx output.docx --preset official
```

从 Markdown 或纯文本生成格式化 DOCX：

```bash
uv run --with python-docx python scripts/from_text.py input.md output.docx --title "工作方案"
```

## 常用参数

```bash
--preset official
--revision
--deep-clean
--smart-table-align
--no-page-number
--page-number-style dash|plain|page_text|page_total
--page-number-position outside|left|center|right
--space-mode remove_all|keep_en_boundary|keep_all
```

## 脚本说明

| 脚本 | 用途 |
| --- | --- |
| `scripts/process.py` | 主入口：`smart`、`analyze`、`punctuation`、`format`。 |
| `scripts/formatter.py` | 公文格式化引擎。 |
| `scripts/punctuation.py` | 标点和空格修复。 |
| `scripts/from_text.py` | 纯文本/Markdown 转 DOCX。 |
| `scripts/analyzer.py` | 格式诊断。 |
| `scripts/converter.py` | Windows `.doc/.wps` 转换辅助。 |

## 注意

- `.docx` 是最稳定的处理格式。
- 自动排版前建议保留原文件备份。

## 许可证

MIT
