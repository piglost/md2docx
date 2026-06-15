# md2docx — Markdown → Word 转换器（法律写作专用）

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-green.svg)](https://www.python.org/)
[![Pandoc 3.0+](https://img.shields.io/badge/Pandoc-3.0%2B-orange.svg)](https://pandoc.org/)

将 Markdown 法学论文/法律文档一键转换为 Word (.docx)，自动应用《中国法学》投稿格式。

## ✨ 核心功能

| 功能 | 说明 |
|------|------|
| 📝 **引用→脚注** | `[1]` `[2-5]` `[3,4]` → Word 脚注，支持范围引用和逗号引用 |
| 🎨 **精准排版** | 自动应用《中国法学》格式（宋体/黑体，各级标题字号，两端对齐） |
| ⬛ **纯黑文字** | 三道防线确保零颜色残留 |
| 📐 **首行缩进** | 正文自动首行缩进 2 字符 |
| 🔤 **字体统一** | 中文宋体/黑体 + 英文 Times New Roman |

## 📋 格式规范

支持两种期刊格式，自动检测或手动指定。

### 《中国法学》（`-f faxue`）

| 元素 | 字体 | 字号 | 加粗 | 对齐 |
|------|------|------|------|------|
| 正标题 | 宋体 | 二号 22pt | ✓ | 居中 |
| 作者信息 | 宋体 | 小三 15pt | — | 居中 |
| 内容提要/关键词 | 宋体 | 小四 12pt | 按原文 | 两端对齐 |
| 一级标题 | 黑体 | 四号 14pt | — | 居中 |
| 二级标题 | 黑体 | 小四 12pt | — | 两端对齐 |
| 三级标题 | 宋体 | 小四 12pt | — | 两端对齐 |
| **正文** | **宋体** | **小四 12pt** | **—** | **两端对齐** |
| 脚注 | 宋体 | 小五 9pt | — | 连续编号 |

### 《中国社会科学》（`-f sheke`）

| 元素 | 字体 | 字号 | 加粗 | 对齐 |
|------|------|------|------|------|
| 正标题 | 宋体 | 二号 22pt | ✓ | 居中 |
| 一级标题 | 宋体 | 四号 14pt | — | 居中 |
| 二级标题 | 宋体 | 小四 12pt | — | 两端对齐+首行缩进 |
| **正文** | **宋体** | **小四 12pt** | **—** | **两端对齐** |
| 脚注 | 仿宋 | 五号 10.5pt | — | 每页重编号 |

### 段落格式

- 正文：两端对齐，首行缩进 2 字符，1.5 倍行距
- 一级标题：居中，段前段后间距
- 所有文字颜色：纯黑 `#000000`

## 🚀 快速开始

### 依赖

```bash
brew install pandoc          # macOS
# 或 apt-get install pandoc  # Linux
```

### 安装

```bash
git clone https://github.com/piglost/md2docx.git ~/.claude/tools/md2docx
```

### 作为 Claude Code / Hermes Agent Skill 使用

复制 `SKILL.md` 到 skills 目录：

```bash
mkdir -p ~/.claude/skills
cp SKILL.md ~/.claude/skills/md2docx.md
```

然后对 AI 说：「把这个 md 转成 word」即可。

### 命令行使用

```bash
python3 md2docx.py 论文.md                    # 生成同名 .docx
python3 md2docx.py 论文.md 输出.docx           # 指定输出文件名
python3 md2docx.py 论文.md --dry-run           # 预览预处理结果
python3 md2docx.py 论文.md --toc               # 插入自动目次（实验性）
python3 md2docx.py 论文.md -f faxue            # 强制《中国法学》格式
python3 md2docx.py 论文.md -f sheke            # 强制《中国社会科学》格式
```

### 自动检测逻辑

`--format`（`-f`）参数支持三种值：
- `auto`（默认）：自动检测——①②③ 或 〔1〕 标记 → 中国社会科学；[1] 标记 → 中国法学
- `faxue`：强制《中国法学》格式（宋体/黑体混用，脚注宋体小五，连续编号）
- `sheke`：强制《中国社会科学》格式（全文宋体，脚注仿宋五号，每页重编号）

## 📂 文件结构要求

```markdown
# 论文标题

**作者：XXX**
**单位：XXX大学法学院**

## 内容提要

内容提要正文段落……

**关键词：** 关键词1；关键词2；关键词3

## 一、一级标题

正文内容，引用使用 [1] 标记。[2][3]

### （一）二级标题

正文内容。更多引用 [4-6]。

## 参考文献

1. 作者A：《文献标题》，出版社2020年版，第10页。
2. 作者B：《文献标题》，载《期刊名》2021年第3期。
```

### 支持的引用格式

| 格式 | 输入 | 输出 |
|------|------|------|
| 单引用 | `[1]` | 脚注 1 |
| 连续多引用 | `[1][2][3]` | 脚注 1、2、3 |
| 范围引用 | `[1-5]` | 脚注 1-5 展开 |
| 逗号引用 | `[1,3,5]` | 脚注 1、3、5 |

### 参考文献条目格式

支持两种格式：
- `[1] 作者：《标题》，出版社，年份。`（方括号）
- `1. 作者：《标题》，出版社，年份。`（数字句号）

## 🔧 转换流水线

```
Markdown (.md)
    │
    ▼
[1] 预处理：引用标记 → Pandoc 脚注语法
    │
    ▼
[2] Pandoc 转换：md → docx + strip-color.lua 去色
    │
    ▼
[3] Heading 修复：移除 Pandoc 蓝绿标题残留
    │
    ▼
[4] 《中国法学》格式：字体/字号/对齐/缩进精确设置
    │
    ▼
Word (.docx) ✅
```

## 📄 文件说明

| 文件 | 作用 |
|------|------|
| `md2docx.py` | 主转换脚本 |
| `zhongguo-faxue-format.py` | 《中国法学》格式处理器 |
| `zhongguo-sheke-format.py` | 《中国社会科学》格式处理器 |
| `strip-color.lua` | Pandoc 颜色剥离过滤器 |
| `SKILL.md` | AI Skill 定义文件 |

## 🛠️ 技术栈

- Python 3.9+ — 预处理 & 后处理
- Pandoc 3.0+ — Markdown → DOCX 核心转换
- Lua — Pandoc AST 过滤器
- OOXML — Word 文档 XML 精确操控

## 📜 License

MIT License. 详见 [LICENSE](LICENSE) 文件。

## 🙏 致谢

- [Pandoc](https://pandoc.org/) — 文档转换引擎
- 《中国法学》编辑部 — 格式规范参照
