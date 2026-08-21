# document-format-skills 公文格式助手

用于清理和格式化中文 Word 文档的命令行 Skill。当前核心处理逻辑已同步到 [Document Format GUI](https://github.com/KaguraNanaga/docformat-gui) v1.8.7，适合 Codex、Claude Code、OpenCode 等 agent 在没有桌面 UI 的场景下直接调用。

英文说明 / English docs: [README.md](./README.md)

## 功能

- 智能一键处理：标点/空格清理 + 文档格式统一。
- 格式诊断：检查标点、序号、段落、字体问题。
- 提供统一的中文公文排版格式：大标题为方正小标宋简体小二号，正文为仿宋_GB2312三号，一级标题为方正黑体_GBK三号，二级标题为楷体_GB2312三号，英文为 Times New Roman；上下页边距25.4mm、左右31.8mm，固定行距30磅，正文首行缩进2字符。
- 页码支持样式、位置、距版心偏移、替换已有页码，并避免覆盖非页码页脚内容。
- 表格格式统一，支持可选的智能对齐。
- 支持输出 Word 修订标记。
- 支持 macOS 常用中文公文字体回退。
- Windows 下可借助 WPS Office 或 Microsoft Word 处理 `.doc` / `.wps`。
- 支持从纯文本或 Markdown 生成并格式化 DOCX。

## 环境要求

- Python 3.8+
- `python-docx`
- Windows `.doc/.wps` 转换需要 `pywin32`、WPS Office 或 Microsoft Word

推荐使用 `uv` 临时安装依赖：

```bash
uv run --with python-docx python scripts/process.py --help
```

Windows 处理 `.doc/.wps` 时：

```bash
uv run --with python-docx --with pywin32 python scripts/process.py --help
```

## 快速开始

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
| 页码数字 | Times New Roman |
| 页码 | 默认添加，可设置样式、位置、偏移和是否替换已有页码 |

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

只应用格式：

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
- `.doc/.wps` 需要 Windows、WPS Office 或 Microsoft Word。
- 自动排版前建议保留原文件备份。

## 许可证

MIT
